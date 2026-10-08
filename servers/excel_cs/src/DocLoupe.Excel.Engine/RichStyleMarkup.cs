using System.Xml;
using DocLoupe.Excel.Package;

namespace DocLoupe.Excel.Engine;

public sealed record RichStylePatch(bool? Bold, bool? Italic, string? Color);

public static class RichStyleMarkup
{
    public static RichSetValue ApplyAll(XmlElement container, RichStylePatch patch)
    {
        if (container.NamespaceURI != PackageStore.Main || container.LocalName is not ("is" or "si"))
            throw new NotSupportedException("rich_style requires inline or shared rich text");
        var elements = ChildElements(container);
        if (elements.Length is < 1 or > 256 || elements.Any(element =>
                element.LocalName != "r" || element.NamespaceURI != PackageStore.Main || element.HasAttributes))
            throw new NotSupportedException("rich_style supports only bounded rich runs without phonetics");
        var runs = elements.Select(ReadRun).Select(run => run with
        {
            Bold = patch.Bold ?? run.Bold,
            Italic = patch.Italic ?? run.Italic,
            Color = patch.Color ?? run.Color
        }).ToArray();
        return RichSetMarkup.Parse(RichSetMarkup.Render(runs));
    }

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
