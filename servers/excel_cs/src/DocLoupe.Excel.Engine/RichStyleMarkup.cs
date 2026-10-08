using System.Globalization;
using System.Text;
using System.Xml;
using DocLoupe.Excel.Package;

namespace DocLoupe.Excel.Engine;

public sealed record RichStylePatch(bool? Bold, bool? Italic, string? Color);
public sealed record RichStyleSpan(int Start, int End);
public sealed record RichStyleMatch(string Text, bool Normalize);

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
        if (match is not null) span = FindSpan(text, graphemes, match);
        if (span is null) throw new InvalidOperationException("Missing rich_style selection");
        if (span.Start < 0 || span.End > graphemes.Length || span.Start >= span.End)
            throw new ArgumentOutOfRangeException(nameof(span), "rich_style range must select existing graphemes");
        var start = graphemes[span.Start];
        var end = span.End == graphemes.Length ? text.Length : graphemes[span.End];
        var styled = new List<RichSetRun>();
        var offset = 0;
        foreach (var run in runs)
        {
            var before = Math.Clamp(start - offset, 0, run.Text.Length);
            var after = Math.Clamp(end - offset, 0, run.Text.Length);
            if (before > 0) styled.Add(run with { Text = run.Text[..before] });
            if (after > before) styled.Add(PatchRun(run with { Text = run.Text[before..after] }, patch));
            if (after < run.Text.Length) styled.Add(run with { Text = run.Text[after..] });
            offset += run.Text.Length;
        }
        return RichSetMarkup.Parse(RichSetMarkup.Render(styled));
    }

    private static RichStyleSpan FindSpan(string text, int[] graphemes, RichStyleMatch match)
    {
        var sought = StringInfo.ParseCombiningCharacters(match.Text);
        if (sought.Length == 0 || sought.Length > graphemes.Length)
            throw new ArgumentException("rich_style match not found");
        var expected = Enumerable.Range(0, sought.Length).Select(index => Normalize(
            match.Text[sought[index]..(index + 1 == sought.Length ? match.Text.Length : sought[index + 1])], match.Normalize)).ToArray();
        var actual = Enumerable.Range(0, graphemes.Length).Select(index => Normalize(
            text[graphemes[index]..(index + 1 == graphemes.Length ? text.Length : graphemes[index + 1])], match.Normalize)).ToArray();
        for (var start = 0; start <= graphemes.Length - expected.Length; start++)
        {
            if (actual.AsSpan(start, expected.Length).SequenceEqual(expected))
                return new RichStyleSpan(start, start + expected.Length);
        }
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
