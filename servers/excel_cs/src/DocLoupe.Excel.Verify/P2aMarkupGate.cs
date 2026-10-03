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
                    var oldData = oldRoot.DocumentElement.ChildNodes.OfType<XmlElement>()
                        .SingleOrDefault(element => element.LocalName == "sheetData" && element.NamespaceURI == Main);
                    var newData = newRoot.DocumentElement!.ChildNodes.OfType<XmlElement>()
                        .SingleOrDefault(element => element.LocalName == "sheetData" && element.NamespaceURI == Main);
                    if (newData is not null)
                    {
                        if (oldData is null ? newData.Prefix != oldRoot.DocumentElement.Prefix : newData.Name != oldData.Name)
                            issues.Add(new("G3", "PREFIX_REWRITTEN", $"{part}: sheetData"));
                        var oldRows = oldData?.ChildNodes.OfType<XmlElement>()
                            .Where(element => element.LocalName == "row" && element.NamespaceURI == Main)
                            .ToDictionary(element => element.GetAttribute("r"))
                            ?? new Dictionary<string, XmlElement>();
                        var rowPrefix = oldData?.ChildNodes.OfType<XmlElement>()
                            .FirstOrDefault(element => element.LocalName == "row" && element.NamespaceURI == Main)?.Prefix
                            ?? oldData?.Prefix ?? oldRoot.DocumentElement.Prefix;
                        foreach (var row in newData.ChildNodes.OfType<XmlElement>()
                            .Where(element => element.LocalName == "row" && element.NamespaceURI == Main))
                        {
                            oldRows.TryGetValue(row.GetAttribute("r"), out var previous);
                            if (previous is null ? row.Prefix != rowPrefix : row.Name != previous.Name)
                                issues.Add(new("G3", "PREFIX_REWRITTEN", $"{part}: row {row.GetAttribute("r")}"));
                        }
                    }
                    var oldCells = oldRoot.GetElementsByTagName("c", Main).OfType<XmlElement>()
                        .ToDictionary(cell => cell.GetAttribute("r"));
                    foreach (var cell in newRoot.GetElementsByTagName("c", Main).OfType<XmlElement>())
                    {
                        oldCells.TryGetValue(cell.GetAttribute("r"), out var previous);
                        if (previous is null ? cell.Prefix != newRoot.DocumentElement!.Prefix : previous.Name != cell.Name)
                            issues.Add(new("G3", "PREFIX_REWRITTEN", $"{part}: {cell.GetAttribute("r")}"));
                        CheckCellPrefixShape(previous, cell, part, cell.GetAttribute("r"), issues);
                    }
                }
            }
            catch (XmlException exception)
            {
                issues.Add(new("G3", "INVALID_XML", $"{part}: {exception.Message}"));
            }
        }
        return issues;
    }

    private static void CheckCellPrefixShape(XmlElement? original, XmlElement written,
        string part, string address, List<GateIssue> issues)
    {
        var oldChildren = original?.ChildNodes.OfType<XmlElement>()
            .Where(child => child.NamespaceURI == Main)
            .GroupBy(child => child.LocalName, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal)
            ?? new Dictionary<string, XmlElement[]>(StringComparer.Ordinal);
        var indices = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var child in written.ChildNodes.OfType<XmlElement>().Where(child => child.NamespaceURI == Main))
        {
            var index = indices.GetValueOrDefault(child.LocalName);
            indices[child.LocalName] = index + 1;
            var previous = oldChildren.TryGetValue(child.LocalName, out var matches) && index < matches.Length
                ? matches[index] : null;
            if (previous is null ? child.Prefix != written.Prefix : child.Name != previous.Name)
                issues.Add(new("G3", "PREFIX_REWRITTEN", $"{part}: {address}"));
            CheckCellPrefixShape(previous, child, part, address, issues);
        }
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
