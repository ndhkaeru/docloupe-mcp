using System.IO.Compression;
using System.Xml;

namespace DocLoupe.Excel.Verify;

public sealed record SheetSummary(string Name, string Part, string State);
public sealed record CellSummary(string Address, string Type, string? RawValue, string? Formula);
public sealed record WorkbookSummary(string Path, IReadOnlyList<SheetSummary> Sheets, IReadOnlyList<CellSummary> FirstSheetCells);
public sealed record VerificationSummary(string Status, IReadOnlyList<MarkupIssue> PackageIssues,
    IReadOnlyList<MarkupIssue> MarkupIssues, IReadOnlyList<string> UnverifiedGates);

public static class WorkbookReader
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypeNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static WorkbookSummary Peek(string path, int maxCells = 24)
    {
        if (maxCells is < 0 or > 2000) throw new ArgumentOutOfRangeException(nameof(maxCells));
        using var archive = ZipFile.OpenRead(path);
        var relationships = ReadWorkbookRelationships(archive);
        var sheets = ReadSheets(archive, relationships);
        var cells = sheets.Count == 0 ? [] : ReadCells(archive, sheets[0].Part, maxCells);
        return new WorkbookSummary(path, sheets, cells);
    }

    public static VerificationSummary VerifyPartial(string path)
    {
        var packageIssues = new List<MarkupIssue>();
        var markupIssues = new List<MarkupIssue>();
        using var archive = ZipFile.OpenRead(path);
        var duplicates = archive.Entries.GroupBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1);
        foreach (var duplicate in duplicates)
            packageIssues.Add(new MarkupIssue("DUPLICATE_ENTRY", duplicate.Key));
        foreach (var required in new[] { "[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels" })
        {
            if (archive.GetEntry(required) is null)
                packageIssues.Add(new MarkupIssue("MISSING_PART", required));
        }

        if (archive.GetEntry("[Content_Types].xml") is { } types)
        {
            try
            {
                using var reader = CreateReader(types);
                reader.MoveToContent();
                if (reader.LocalName != "Types" || reader.NamespaceURI != ContentTypeNamespace)
                    packageIssues.Add(new MarkupIssue("INVALID_CONTENT_TYPES_ROOT", types.FullName));
            }
            catch (XmlException exception)
            {
                packageIssues.Add(new MarkupIssue("INVALID_CONTENT_TYPES_XML", exception.Message));
            }
        }

        foreach (var entry in archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            using var stream = entry.Open();
            foreach (var issue in MarkupCompatibilityVerifier.Check(stream).Issues)
                markupIssues.Add(new MarkupIssue(issue.Code, $"{entry.FullName}: {issue.Detail}"));
        }

        var status = packageIssues.Count + markupIssues.Count > 0 ? "failed" : "unverified";
        return new VerificationSummary(status, packageIssues, markupIssues, ["G1_REMAINING", "G2", "G4", "G5", "G6", "G7"]);
    }

    private static Dictionary<string, string> ReadWorkbookRelationships(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new InvalidDataException("Missing workbook relationships");
        var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
        using var reader = CreateReader(entry);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship"
                || reader.NamespaceURI != PackageRelationshipNamespace)
                continue;
            var id = reader.GetAttribute("Id");
            var target = reader.GetAttribute("Target");
            if (id is not null && target is not null && reader.GetAttribute("TargetMode") != "External")
                relationships[id] = ResolvePartPath(target);
        }
        return relationships;
    }

    private static List<SheetSummary> ReadSheets(ZipArchive archive, Dictionary<string, string> relationships)
    {
        var entry = archive.GetEntry("xl/workbook.xml") ?? throw new InvalidDataException("Missing workbook part");
        var sheets = new List<SheetSummary>();
        using var reader = CreateReader(entry);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "sheet"
                || reader.NamespaceURI != SpreadsheetNamespace)
                continue;
            var id = reader.GetAttribute("id", RelationshipNamespace);
            if (id is null || !relationships.TryGetValue(id, out var part))
                throw new InvalidDataException($"Sheet has unresolved relationship {id}");
            sheets.Add(new SheetSummary(reader.GetAttribute("name") ?? "", part, reader.GetAttribute("state") ?? "visible"));
        }
        return sheets;
    }

    private static List<CellSummary> ReadCells(ZipArchive archive, string part, int maxCells)
    {
        var entry = archive.GetEntry(part) ?? throw new InvalidDataException($"Missing sheet part {part}");
        var cells = new List<CellSummary>();
        using var reader = CreateReader(entry);
        while (cells.Count < maxCells && reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c"
                || reader.NamespaceURI != SpreadsheetNamespace)
                continue;
            var address = reader.GetAttribute("r") ?? "";
            var type = reader.GetAttribute("t") ?? "n";
            string? value = null;
            string? formula = null;
            var text = new System.Text.StringBuilder();
            using (var cell = reader.ReadSubtree())
            {
                // ReadElementContentAsString and Skip already advance to the next node, so the loop
                // only calls Read() when nothing consumed the current node. Otherwise a sibling
                // directly after <f> (the cached <v>) would be skipped.
                cell.Read();
                while (!cell.EOF)
                {
                    if (cell.NodeType == XmlNodeType.Element && cell.NamespaceURI == SpreadsheetNamespace)
                    {
                        switch (cell.LocalName)
                        {
                            case "rPh":
                            case "phoneticPr":
                                cell.Skip(); // phonetic guide text is not part of the cell value
                                continue;
                            case "f":
                                formula = cell.ReadElementContentAsString();
                                continue;
                            case "v":
                                value = cell.ReadElementContentAsString();
                                continue;
                            case "t":
                                text.Append(cell.ReadElementContentAsString());
                                continue;
                        }
                    }
                    cell.Read();
                }
            }
            if (value is null && text.Length > 0) value = text.ToString();
            cells.Add(new CellSummary(address, type, value, formula));
        }
        return cells;
    }

    private static string ResolvePartPath(string target)
    {
        var segments = new List<string>();
        if (!target.StartsWith('/')) segments.Add("xl");
        foreach (var segment in target.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".") continue;
            if (segment == "..")
            {
                if (segments.Count == 0) throw new InvalidDataException("Relationship escapes the package");
                segments.RemoveAt(segments.Count - 1);
            }
            else segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    private static XmlReader CreateReader(ZipArchiveEntry entry)
    {
        var stream = entry.Open();
        return XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            CloseInput = true
        });
    }
}
