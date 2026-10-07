using System.Text.RegularExpressions;
using System.Xml;
using DocLoupe.Excel.Package;

namespace DocLoupe.Excel.Engine;

public sealed record RichSetRun(string Text, bool? Bold, bool? Italic, string? Color);
public sealed record RichSetValue(string Text, IReadOnlyList<RichSetRun> Runs);

public static class RichSetMarkup
{
    private static readonly Regex ColorPattern = new(@"\A(?:[A-Fa-f0-9]{6}|[A-Fa-f0-9]{8})\z",
        RegexOptions.CultureInvariant | RegexOptions.NonBacktracking);

    public static RichSetValue Parse(string markup)
    {
        if (markup.Length > 8192) throw new NotSupportedException("rich_set exceeds 8192 characters");
        var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
        using var reader = XmlReader.Create(new StringReader("<rich>" + markup + "</rich>"),
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        document.Load(reader);
        var nodes = document.DocumentElement!.ChildNodes;
        if (nodes.Count is < 1 or > 256) throw new NotSupportedException("rich_set requires 1..256 runs");
        var runs = new List<RichSetRun>(nodes.Count);
        foreach (XmlNode node in nodes)
        {
            if (node is not XmlElement { Name: "r", NamespaceURI: "" } run ||
                run.ChildNodes.Count != 1 || run.FirstChild!.NodeType is not
                    (XmlNodeType.Text or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace))
                throw new NotSupportedException("rich_set supports text-only r elements without outer text");
            bool? bold = null;
            bool? italic = null;
            string? color = null;
            foreach (XmlAttribute attribute in run.Attributes)
            {
                if (attribute.NamespaceURI.Length != 0) throw new NotSupportedException("rich_set does not support namespaces");
                switch (attribute.Name)
                {
                    case "b": bold = Boolean(attribute.Value); break;
                    case "i": italic = Boolean(attribute.Value); break;
                    case "color" when ColorPattern.IsMatch(attribute.Value):
                        color = (attribute.Value.Length == 6 ? "FF" : "") + attribute.Value.ToUpperInvariant();
                        break;
                    default: throw new NotSupportedException("Unsupported rich_set run attribute: " + attribute.Name);
                }
            }
            runs.Add(new RichSetRun(run.InnerText, bold, italic, color));
        }
        return new RichSetValue(string.Concat(runs.Select(run => run.Text)), runs);
    }

    public static void Append(XmlDocument document, XmlElement inline, RichSetValue value)
    {
        foreach (var run in value.Runs)
        {
            var element = document.CreateElement(inline.Prefix, "r", PackageStore.Main);
            if (run.Bold is not null || run.Italic is not null || run.Color is not null)
            {
                var properties = document.CreateElement(inline.Prefix, "rPr", PackageStore.Main);
                foreach (var (name, enabled) in new[] { ("b", run.Bold), ("i", run.Italic) })
                {
                    if (enabled is null) continue;
                    var property = document.CreateElement(inline.Prefix, name, PackageStore.Main);
                    property.SetAttribute("val", enabled.Value ? "1" : "0");
                    properties.AppendChild(property);
                }
                if (run.Color is { } rgb)
                {
                    var color = document.CreateElement(inline.Prefix, "color", PackageStore.Main);
                    color.SetAttribute("rgb", rgb);
                    properties.AppendChild(color);
                }
                element.AppendChild(properties);
            }
            var text = document.CreateElement(inline.Prefix, "t", PackageStore.Main);
            text.InnerText = run.Text;
            if (run.Text.Length > 0 && (char.IsWhiteSpace(run.Text[0]) || char.IsWhiteSpace(run.Text[^1])))
                text.SetAttribute("xml:space", "preserve");
            element.AppendChild(text);
            inline.AppendChild(element);
        }
    }

    private static bool Boolean(string value) => value switch
    {
        "true" or "1" => true,
        "false" or "0" => false,
        _ => throw new NotSupportedException("rich_set boolean run attributes require true, false, 1 or 0")
    };
}
