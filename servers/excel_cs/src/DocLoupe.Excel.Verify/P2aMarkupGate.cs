using System.IO.Compression;
using System.Text;
using System.Xml;

namespace DocLoupe.Excel.Verify;

public static class P2aMarkupGate
{
    private static readonly UTF8Encoding Utf8 = new(false, true);
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static IReadOnlyList<GateIssue> Check(string source, string written, IEnumerable<string> editedParts)
    {
        var issues = new List<GateIssue>();
        using var original = ZipFile.OpenRead(source);
        using var output = ZipFile.OpenRead(written);
        foreach (var part in editedParts.Where(part => part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            var after = output.Entries.SingleOrDefault(entry => entry.FullName.Equals(part, StringComparison.OrdinalIgnoreCase));
            if (after is null) continue;
            byte[] writtenBytes = Read(after);
            foreach (var issue in MarkupCompatibilityVerifier.Check(new MemoryStream(writtenBytes)).Issues)
                issues.Add(new GateIssue("G3", issue.Code, $"{part}: {issue.Detail}"));
            var before = original.Entries.SingleOrDefault(entry => entry.FullName.Equals(part, StringComparison.OrdinalIgnoreCase));
            if (before is null) continue;
            byte[] originalBytes = Read(before);
            if (Declaration(originalBytes) != Declaration(writtenBytes))
                issues.Add(new("G3", "DECLARATION_CHANGED", part));
            try
            {
                var oldRoot = Parse(originalBytes);
                var newRoot = Parse(writtenBytes);
                if (oldRoot.DocumentElement?.Name != newRoot.DocumentElement?.Name)
                    issues.Add(new("G3", "PREFIX_REWRITTEN", part));
                var oldNamespaces = Namespaces(oldRoot.DocumentElement!);
                var newNamespaces = Namespaces(newRoot.DocumentElement!);
                if (!oldNamespaces.SequenceEqual(newNamespaces)) issues.Add(new("G3", "ROOT_NAMESPACE_CHANGED", part));
                if (oldRoot.DocumentElement!.NamespaceURI == Main)
                {
                    var cells = newRoot.GetElementsByTagName("c", Main).OfType<XmlElement>()
                        .ToDictionary(cell => cell.GetAttribute("r"));
                    foreach (var cell in oldRoot.GetElementsByTagName("c", Main).OfType<XmlElement>())
                        if (cells.TryGetValue(cell.GetAttribute("r"), out var edited) && cell.Name != edited.Name)
                            issues.Add(new("G3", "PREFIX_REWRITTEN", $"{part}: {cell.GetAttribute("r")}"));
                }
            }
            catch (XmlException exception)
            {
                issues.Add(new("G3", "INVALID_XML", $"{part}: {exception.Message}"));
            }
        }
        return issues;
    }

    private static XmlDocument Parse(byte[] content)
    {
        var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
        using var reader = XmlReader.Create(new MemoryStream(content), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        document.Load(reader);
        return document;
    }

    private static byte[] Read(ZipArchiveEntry entry)
    {
        using var input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static string Declaration(byte[] content)
    {
        var xml = Utf8.GetString(content).TrimStart('\uFEFF');
        return xml.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) ? xml[..(xml.IndexOf("?>", StringComparison.Ordinal) + 2)] : "";
    }

    private static IReadOnlyList<string> Namespaces(XmlElement root) => root.Attributes.OfType<XmlAttribute>()
        .Where(attribute => attribute.Name == "xmlns" || attribute.Prefix == "xmlns")
        .Select(attribute => attribute.Name + "=" + attribute.Value).OrderBy(value => value, StringComparer.Ordinal).ToArray();
}
