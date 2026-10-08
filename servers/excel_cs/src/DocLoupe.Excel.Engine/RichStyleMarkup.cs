using System.Globalization;
using System.Text;
using System.Xml;
using DocLoupe.Excel.Package;

namespace DocLoupe.Excel.Engine;

public sealed record RichStylePatch(bool? Bold, bool? Italic, string? Color);
public sealed record RichStyleSpan(int Start, int End);
public sealed record RichStyleMatch(string Text, bool Normalize, int? Occurrence = 1);

public static class RichStyleMarkup
{
    private static readonly bool HasUnicodeNfc = "e\u0323\u0302".Normalize(NormalizationForm.FormC) == "ệ";

    public static RichSetValue Apply(XmlElement container, RichStylePatch patch, RichStyleSpan? span, RichStyleMatch? match = null)
    {
        if (container.NamespaceURI != PackageStore.Main || container.LocalName is not ("is" or "si"))
            throw new NotSupportedException("rich_style requires inline or shared rich text");
        var elements = ChildElements(container);
        if (elements.Length is < 1 or > 256 || elements.Any(element =>
                element.LocalName != "r" || element.NamespaceURI != PackageStore.Main || element.HasAttributes))
            throw new NotSupportedException("rich_style supports only bounded rich runs without phonetics");
        var runs = elements.Select(ReadRun).ToArray();
        if (span is null && match is null)
            return RichSetMarkup.Parse(RichSetMarkup.Render(runs.Select(run => PatchRun(run, patch)).ToArray()));
        var text = string.Concat(runs.Select(run => run.Text));
        if (text.Length > 8192) throw new NotSupportedException("rich_style exceeds 8192 characters");
        var graphemes = StringInfo.ParseCombiningCharacters(text);
        var selections = match is null ? [span ?? throw new InvalidOperationException("Missing rich_style selection")]
            : FindSpans(text, graphemes, match);
        var merged = new List<RichStyleSpan>();
        foreach (var selection in selections)
        {
            if (selection.Start < 0 || selection.End > graphemes.Length || selection.Start >= selection.End)
                throw new ArgumentOutOfRangeException(nameof(span), "rich_style range must select existing graphemes");
            if (merged.Count > 0 && merged[^1].End == selection.Start)
                merged[^1] = merged[^1] with { End = selection.End };
            else merged.Add(selection);
        }
        var offsets = merged.Select(selection => (Start: graphemes[selection.Start],
            End: selection.End == graphemes.Length ? text.Length : graphemes[selection.End])).ToArray();
        var styled = new List<RichSetRun>();
        var offset = 0;
        foreach (var run in runs)
        {
            var runEnd = offset + run.Text.Length;
            var position = offset;
            foreach (var selection in offsets)
            {
                if (selection.End <= position) continue;
                if (selection.Start >= runEnd) break;
                var beforeEnd = Math.Max(position, selection.Start);
                if (beforeEnd > position)
                    styled.Add(run with { Text = run.Text[(position - offset)..(beforeEnd - offset)] });
                var selectedEnd = Math.Min(runEnd, selection.End);
                styled.Add(PatchRun(run with { Text = run.Text[(beforeEnd - offset)..(selectedEnd - offset)] }, patch));
                position = selectedEnd;
            }
            if (position < runEnd)
                styled.Add(run with { Text = run.Text[(position - offset)..] });
            offset = runEnd;
        }
        return RichSetMarkup.Parse(RichSetMarkup.Render(styled));
    }

    private static RichStyleSpan[] FindSpans(string text, int[] graphemes, RichStyleMatch match)
    {
        var sought = StringInfo.ParseCombiningCharacters(match.Text);
        if (sought.Length == 0 || sought.Length > graphemes.Length)
            throw new ArgumentException("rich_style match not found");
        var expected = Enumerable.Range(0, sought.Length).Select(index => Normalize(
            match.Text[sought[index]..(index + 1 == sought.Length ? match.Text.Length : sought[index + 1])], match.Normalize)).ToArray();
        var actual = Enumerable.Range(0, graphemes.Length).Select(index => Normalize(
            text[graphemes[index]..(index + 1 == graphemes.Length ? text.Length : graphemes[index + 1])], match.Normalize)).ToArray();
        var found = new List<RichStyleSpan>();
        for (var start = 0; start <= graphemes.Length - expected.Length;)
        {
            if (actual.AsSpan(start, expected.Length).SequenceEqual(expected))
            {
                found.Add(new RichStyleSpan(start, start + expected.Length));
                if (found.Count == match.Occurrence) return [found[^1]];
                start += expected.Length;
            }
            else start++;
        }
        if (match.Occurrence is null && found.Count > 0) return found.ToArray();
        throw new ArgumentException("rich_style match not found");
    }

    private static string Normalize(string value, bool nfc)
    {
        if (!nfc) return value;
        if (!HasUnicodeNfc && value.Any(character => character > 0x7f))
            throw new NotSupportedException("rich_style NFC requires Unicode normalization data");
        return value.Normalize(NormalizationForm.FormC);
    }

    private static RichSetRun PatchRun(RichSetRun run, RichStylePatch patch) => run with
    {
        Bold = patch.Bold ?? run.Bold,
        Italic = patch.Italic ?? run.Italic,
        Color = patch.Color ?? run.Color
    };

    private static RichSetRun ReadRun(XmlElement element)
    {
        var children = ChildElements(element);
        var properties = children.FirstOrDefault(child => child.LocalName == "rPr" && child.NamespaceURI == PackageStore.Main);
        var text = children.LastOrDefault();
        if (children.Length != (properties is null ? 1 : 2) ||
            properties is not null && children[0] != properties ||
            text is null || text.LocalName != "t" || text.NamespaceURI != PackageStore.Main ||
            text.Attributes.OfType<XmlAttribute>().Any(attribute =>
                attribute.LocalName != "space" || attribute.NamespaceURI != "http://www.w3.org/XML/1998/namespace" ||
                attribute.Value is not ("preserve" or "default")) ||
            text.ChildNodes.Count != 1 || text.FirstChild!.NodeType is not
                (XmlNodeType.Text or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace))
            throw new NotSupportedException("rich_style cannot model this run markup");
        bool? bold = null;
        bool? italic = null;
        string? color = null;
        if (properties is not null)
        {
            if (properties.HasAttributes)
                throw new NotSupportedException("rich_style cannot model run properties");
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in ChildElements(properties))
            {
                if (property.NamespaceURI != PackageStore.Main || !seen.Add(property.LocalName) ||
                    property.ChildNodes.Count != 0)
                    throw new NotSupportedException("rich_style cannot model run properties");
                switch (property.LocalName)
                {
                    case "b": bold = ReadFlag(property); break;
                    case "i": italic = ReadFlag(property); break;
                    case "color" when property.Attributes.Count == 1 && property.HasAttribute("rgb"):
                        color = RichSetMarkup.Parse(RichSetMarkup.Render(
                            [new RichSetRun("x", null, null, property.GetAttribute("rgb"))])).Runs[0].Color;
                        break;
                    default: throw new NotSupportedException("rich_style cannot model run properties");
                }
            }
        }
        return new RichSetRun(text.InnerText, bold, italic, color);
    }

    private static bool ReadFlag(XmlElement property)
    {
        if (property.Attributes.Count == 0) return true;
        if (property.Attributes.Count != 1 || !property.HasAttribute("val"))
            throw new NotSupportedException("rich_style cannot model run flag");
        return property.GetAttribute("val") switch
        {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => throw new NotSupportedException("rich_style cannot model run flag")
        };
    }

    private static XmlElement[] ChildElements(XmlElement parent)
    {
        if (parent.ChildNodes.Cast<XmlNode>().Any(node => node is not XmlElement && node.NodeType != XmlNodeType.Whitespace))
            throw new NotSupportedException("rich_style cannot model text or annotations outside rich runs");
        return parent.ChildNodes.OfType<XmlElement>().ToArray();
    }
}
