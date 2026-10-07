using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;

namespace DocLoupe.Excel.Verify;

public static class RichTextAssertions
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Xml = "http://www.w3.org/XML/1998/namespace";
    private static readonly Dictionary<string, string> PropertyNames = new(StringComparer.Ordinal)
    {
        ["b"] = "b", ["i"] = "i", ["s"] = "strike", ["outline"] = "outline",
        ["shadow"] = "shadow", ["condense"] = "condense", ["extend"] = "extend",
        ["u"] = "u", ["sz"] = "sz", ["font"] = "rFont", ["color"] = "color",
        ["va"] = "vertAlign", ["family"] = "family", ["charset"] = "charset", ["scheme"] = "scheme"
    };
    private static readonly Dictionary<string, string> ReverseNames = PropertyNames
        .ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.Ordinal);

    private sealed record Run(bool IsRun, string Text, SortedDictionary<string, string> Properties);

    public static void Validate(string markup) => ParseExpected(markup);

    public static bool Matches(CellRead? cell, string markup)
    {
        if (cell is null || cell.Kind is not ("text" or "inline") || cell.Formula is not null) return false;
        try
        {
            var expected = ParseExpected(markup);
            var document = Load(cell.Kind == "text" ? cell.SharedMarkup : cell.CellMarkup);
            var root = document.DocumentElement;
            var container = cell.Kind == "text" ? root : root?.ChildNodes.OfType<XmlElement>()
                .SingleOrDefault(child => child.LocalName == "is" && child.NamespaceURI == Main);
            if (container is null || container.LocalName is not ("is" or "si") || container.NamespaceURI != Main)
                return false;
            var actual = new List<Run>();
            foreach (var child in container.ChildNodes.OfType<XmlElement>())
            {
                if (child.NamespaceURI != Main) return false;
                if (child.LocalName is "rPh" or "phoneticPr") continue;
                if (child.LocalName == "t") actual.Add(new Run(false, ReadText(child), new SortedDictionary<string, string>(StringComparer.Ordinal)));
                else if (child.LocalName == "r") actual.Add(ReadRun(child));
                else return false;
            }
            return expected == JsonSerializer.Serialize(actual);
        }
        catch (Exception error) when (error is XmlException or InvalidDataException or InvalidOperationException or
                                      NotSupportedException or ArgumentException)
        {
            return false;
        }
    }

    private static string ParseExpected(string markup)
    {
        if (markup.Length > 8192) throw new NotSupportedException("Rich assertion exceeds 8192 characters");
        XmlDocument document;
        try { document = Load("<rich>" + ExpandBooleanAttributes(markup) + "</rich>"); }
        catch (XmlException error) { throw new ArgumentException("Invalid rich markup", nameof(markup), error); }
        var runs = new List<Run>();
        foreach (XmlNode node in document.DocumentElement!.ChildNodes)
        {
            if (runs.Count >= 256) throw new NotSupportedException("Rich assertion exceeds 256 segments");
            if (node.NodeType is XmlNodeType.Text or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace)
                runs.Add(new Run(false, node.Value!, new SortedDictionary<string, string>(StringComparer.Ordinal)));
            else if (node is XmlElement { Name: "r", NamespaceURI: "" } run)
            {
                var attributes = new SortedDictionary<string, string>(StringComparer.Ordinal);
                foreach (XmlAttribute attribute in run.Attributes)
                {
                    if (!PropertyNames.ContainsKey(attribute.Name))
                        throw new NotSupportedException("Unsupported rich attribute: " + attribute.Name);
                    attributes.Add(attribute.Name, Normalize(attribute.Name, attribute.Value));
                }
                if (run.ChildNodes.Count > 1 || run.ChildNodes.Count == 1 &&
                    run.FirstChild!.NodeType is not (XmlNodeType.Text or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace))
                    throw new NotSupportedException("Rich runs may contain text only");
                runs.Add(new Run(true, run.InnerText, attributes));
            }
            else throw new NotSupportedException("Rich markup supports plain text and r elements only");
        }
        return JsonSerializer.Serialize(runs);
    }

    private static string ExpandBooleanAttributes(string markup) => Regex.Replace(markup,
        @"<r(?<attributes>(?:[^"">]|""[^""]*"")*)>", match =>
        {
            var raw = match.Groups["attributes"].Value;
            var result = new StringBuilder("<r");
            var offset = 0;
            foreach (Match attribute in Regex.Matches(raw, @"\s+(?<name>[A-Za-z]+)(?:=""[^""]*"")?"))
            {
                if (attribute.Index != offset) return match.Value;
                result.Append(attribute.Value);
                if (!attribute.Value.Contains('='))
                {
                    if (attribute.Groups["name"].Value == "u") result.Append("=\"single\"");
                    else if (attribute.Groups["name"].Value is "b" or "i" or "s" or "outline" or "shadow" or "condense" or "extend")
                        result.Append("=\"true\"");
                }
                offset += attribute.Length;
            }
            if (!string.IsNullOrWhiteSpace(raw[offset..])) return match.Value;
            return result.Append(raw[offset..]).Append('>').ToString();
        });

    private static Run ReadRun(XmlElement run)
    {
        if (run.Attributes.Count != 0 || run.ChildNodes.OfType<XmlCharacterData>()
            .Any(node => node.NodeType is XmlNodeType.Text or XmlNodeType.CDATA &&
                !string.IsNullOrWhiteSpace(node.Value)))
            throw new NotSupportedException("Unsupported rich run content");
        var properties = new SortedDictionary<string, string>(StringComparer.Ordinal);
        XmlElement? text = null;
        var sawProperties = false;
        foreach (var child in run.ChildNodes.OfType<XmlElement>())
        {
            if (child.NamespaceURI != Main) throw new NotSupportedException("Unsupported rich run namespace");
            if (child.LocalName == "t")
            {
                if (text is not null) throw new NotSupportedException("Multiple rich run texts");
                text = child;
            }
            else if (child.LocalName == "rPr")
            {
                if (sawProperties || text is not null || child.Attributes.Count != 0)
                    throw new NotSupportedException("Invalid rich run properties");
                sawProperties = true;
                foreach (var property in child.ChildNodes.OfType<XmlElement>())
                {
                    if (property.NamespaceURI != Main || !ReverseNames.TryGetValue(property.LocalName, out var name) ||
                        properties.ContainsKey(name) || property.HasChildNodes ||
                        name != "color" && (property.Attributes.Count > 1 ||
                            property.Attributes.Count == 1 && !property.HasAttribute("val")))
                        throw new NotSupportedException("Unsupported rich run property");
                    properties.Add(name, name == "color" ? ReadColor(property) :
                        Normalize(name, property.HasAttribute("val") ? property.GetAttribute("val") : ""));
                }
            }
            else throw new NotSupportedException("Unsupported rich run child");
        }
        if (text is null) throw new NotSupportedException("Rich run has no text");
        return new Run(true, ReadText(text), properties);
    }

    private static string ReadText(XmlElement text)
    {
        if (text.ChildNodes.OfType<XmlElement>().Any() || text.Attributes.OfType<XmlAttribute>()
            .Any(attribute => attribute.LocalName != "space" || attribute.NamespaceURI != Xml ||
                attribute.Value is not ("default" or "preserve")))
            throw new NotSupportedException("Unsupported rich text node");
        return text.InnerText;
    }

    private static string ReadColor(XmlElement color)
    {
        if (color.Attributes.Count == 1)
        {
            var attribute = color.Attributes[0]!;
            if (attribute.Name == "rgb") return Normalize("color", attribute.Value);
            if (attribute.Name == "indexed") return Normalize("color", "indexed:" + attribute.Value);
            if (attribute.Name == "auto" && attribute.Value is "1" or "true")
                return Normalize("color", "auto");
        }
        if (color.HasAttribute("theme") && color.Attributes.Count is 1 or 2 &&
            (color.Attributes.Count == 1 || color.HasAttribute("tint")))
            return Normalize("color", "theme:" + color.GetAttribute("theme") +
                (color.HasAttribute("tint") ? "," + color.GetAttribute("tint") : ""));
        throw new NotSupportedException("Unsupported rich color");
    }

    private static string Normalize(string name, string value)
    {
        if (name is "b" or "i" or "s" or "outline" or "shadow" or "condense" or "extend")
            return value switch
            {
                "" or "1" or "true" => "true",
                "0" or "false" => "false",
                _ => throw new NotSupportedException("Invalid rich Boolean: " + name)
            };
        if (name == "u") return value switch
        {
            "" or "single" => "single",
            "double" or "singleAccounting" or "doubleAccounting" => value,
            _ => throw new NotSupportedException("Invalid underline")
        };
        if (name == "va") return value is "superscript" or "subscript" or "baseline"
            ? value : throw new NotSupportedException("Invalid vertical alignment");
        if (name == "scheme") return value is "minor" or "major" or "none"
            ? value : throw new NotSupportedException("Invalid font scheme");
        if (name is "font") return value.Length is > 0 and <= 255 ? value
            : throw new NotSupportedException("Invalid font name");
        if (name is "family" or "charset") return int.TryParse(value, NumberStyles.None,
            CultureInfo.InvariantCulture, out var number) && number is >= 0 and <= 255
            ? number.ToString(CultureInfo.InvariantCulture) : throw new NotSupportedException("Invalid rich font number");
        if (name == "sz") return decimal.TryParse(value, NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture, out var size) && size > 0 && size <= 409
            ? size.ToString("G29", CultureInfo.InvariantCulture) : throw new NotSupportedException("Invalid rich font size");
        if (name == "color")
        {
            if (value.Length is 6 or 8 && value.All(Uri.IsHexDigit))
                return (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant();
            if (value == "auto") return value;
            var parts = value.Split(':', 2);
            if (parts.Length != 2) throw new NotSupportedException("Invalid rich color");
            var tokens = parts[1].Split(',', 2);
            if (!int.TryParse(tokens[0], NumberStyles.None, CultureInfo.InvariantCulture, out var index) || index < 0)
                throw new NotSupportedException("Invalid rich color index");
            if (parts[0] == "indexed" && tokens.Length == 1) return "indexed:" + index;
            if (parts[0] == "theme" && tokens.Length == 1) return "theme:" + index;
            if (parts[0] == "theme" && tokens.Length == 2 && decimal.TryParse(tokens[1],
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture,
                out var tint) && tint is >= -1 and <= 1)
                return "theme:" + index + "," + tint.ToString("G29", CultureInfo.InvariantCulture);
            throw new NotSupportedException("Invalid rich color");
        }
        throw new NotSupportedException("Unsupported rich attribute: " + name);
    }

    private static XmlDocument Load(string? xml)
    {
        if (xml is null || xml.Length > 1_000_000) throw new InvalidDataException("Rich XML is missing or too large");
        var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
        using var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1_000_000
        });
        document.Load(reader);
        return document;
    }
}
