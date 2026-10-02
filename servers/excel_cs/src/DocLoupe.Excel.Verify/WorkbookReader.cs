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
        if (archive.Entries.GroupBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase).Any(group => group.Count() > 1))
            throw new InvalidDataException("Duplicate OPC part");
        var workbookPart = LocateWorkbookPart(archive);
        var relationships = ReadWorkbookRelationships(archive, workbookPart);
        var sheets = ReadSheets(archive, workbookPart, relationships);
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
        foreach (var required in new[] { "[Content_Types].xml", "_rels/.rels" })
        {
            if (FindEntry(archive, required) is null)
                packageIssues.Add(new MarkupIssue("MISSING_PART", required));
        }

        if (FindEntry(archive, "_rels/.rels") is not null)
        {
            try
            {
                var workbookPart = LocateWorkbookPart(archive);
                var workbookEntry = FindEntry(archive, workbookPart);
                var relationshipsEntry = FindEntry(archive, RelationshipPart(workbookPart));
                if (workbookEntry is null)
                    packageIssues.Add(new MarkupIssue("MISSING_PART", workbookPart));
                if (relationshipsEntry is null)
                    packageIssues.Add(new MarkupIssue("MISSING_PART", RelationshipPart(workbookPart)));
                if (workbookEntry is not null && relationshipsEntry is not null)
                {
                    try
                    {
                        var relationships = ReadWorkbookRelationships(archive, workbookPart);
                        var sheets = ReadSheets(archive, workbookPart, relationships);
                        foreach (var sheet in sheets)
                        {
                            if (FindEntry(archive, sheet.Part) is null)
                                packageIssues.Add(new MarkupIssue("MISSING_PART", sheet.Part));
                        }
                    }
                    catch (XmlException exception)
                    {
                        packageIssues.Add(new MarkupIssue("INVALID_WORKBOOK_STRUCTURE", exception.Message));
                    }
                    catch (InvalidDataException exception)
                    {
                        packageIssues.Add(new MarkupIssue("INVALID_WORKBOOK_STRUCTURE", exception.Message));
                    }
                }
            }
            catch (XmlException exception)
            {
                packageIssues.Add(new MarkupIssue("INVALID_ROOT_RELATIONSHIPS", exception.Message));
            }
            catch (InvalidDataException exception)
            {
                packageIssues.Add(new MarkupIssue("INVALID_ROOT_RELATIONSHIPS", exception.Message));
            }
        }

        if (FindEntry(archive, "[Content_Types].xml") is { } types)
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

    private static string LocateWorkbookPart(ZipArchive archive)
    {
        var entry = FindEntry(archive, "_rels/.rels") ?? throw new InvalidDataException("Missing OPC root relationships");
        using var reader = CreateReader(entry);
        reader.MoveToContent();
        if (reader.LocalName != "Relationships" || reader.NamespaceURI != PackageRelationshipNamespace)
            throw new InvalidDataException("Invalid OPC root relationships");
        string? target = null;
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship"
                || reader.NamespaceURI != PackageRelationshipNamespace
                || reader.GetAttribute("Type") != RelationshipNamespace + "/officeDocument")
                continue;
            if (reader.GetAttribute("TargetMode") == "External")
                throw new InvalidDataException("External officeDocument relationship");
            if (target is not null) throw new InvalidDataException("Multiple officeDocument relationships");
            target = reader.GetAttribute("Target") ?? throw new InvalidDataException("Missing officeDocument target");
        }
        if (target is null) throw new InvalidDataException("Missing officeDocument relationship");
        var part = ResolvePartPath("", target);
        return FindEntry(archive, part)?.FullName ?? part;
    }

    private static Dictionary<string, string> ReadWorkbookRelationships(ZipArchive archive, string workbookPart)
    {
        var entry = FindEntry(archive, RelationshipPart(workbookPart))
            ?? throw new InvalidDataException("Missing workbook relationships");
        var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
        using var reader = CreateReader(entry);
        reader.MoveToContent();
        if (reader.LocalName != "Relationships" || reader.NamespaceURI != PackageRelationshipNamespace)
            throw new InvalidDataException("Invalid workbook relationships");
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship"
                || reader.NamespaceURI != PackageRelationshipNamespace)
                continue;
            var id = reader.GetAttribute("Id");
            var target = reader.GetAttribute("Target");
            if (id is not null && target is not null && reader.GetAttribute("TargetMode") != "External"
                && reader.GetAttribute("Type") == RelationshipNamespace + "/worksheet")
            {
                var part = ResolvePartPath(workbookPart, target);
                if (!relationships.TryAdd(id, FindEntry(archive, part)?.FullName ?? part))
                    throw new InvalidDataException($"Duplicate worksheet relationship: {id}");
            }
        }
        return relationships;
    }

    private static List<SheetSummary> ReadSheets(ZipArchive archive, string workbookPart, Dictionary<string, string> relationships)
    {
        var entry = FindEntry(archive, workbookPart) ?? throw new InvalidDataException("Missing workbook part");
        var sheets = new List<SheetSummary>();
        using var reader = CreateReader(entry);
        reader.MoveToContent();
        if (reader.LocalName != "workbook" || reader.NamespaceURI != SpreadsheetNamespace)
            throw new InvalidDataException("Invalid workbook root");
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
        var entry = FindEntry(archive, part) ?? throw new InvalidDataException($"Missing sheet part {part}");
        var cells = new List<CellSummary>();
        using var reader = CreateReader(entry);
        reader.MoveToContent();
        if (reader.LocalName != "worksheet" || reader.NamespaceURI != SpreadsheetNamespace)
            throw new InvalidDataException("Invalid worksheet root");
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

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string part) => archive.Entries
        .FirstOrDefault(entry => string.Equals(entry.FullName, part, StringComparison.OrdinalIgnoreCase));

    private static string RelationshipPart(string source)
    {
        var slash = source.LastIndexOf('/');
        return (slash < 0 ? "" : source[..(slash + 1)]) + "_rels/" + source[(slash + 1)..] + ".rels";
    }

    private static string ResolvePartPath(string source, string target)
    {
        var clean = Uri.UnescapeDataString(target.Split('#')[0]);
        if (clean.Contains('\\')) throw new InvalidDataException("Backslash in OPC target");
        var combined = clean.StartsWith('/') ? clean[1..] : (source.Contains('/') ? source[..(source.LastIndexOf('/') + 1)] : "") + clean;
        var segments = new List<string>();
        foreach (var segment in combined.Split('/'))
        {
            if (segment == ".") continue;
            if (segment == "..")
            {
                if (segments.Count == 0) throw new InvalidDataException("Relationship escapes the package");
                segments.RemoveAt(segments.Count - 1);
            }
            else if (segment.Length == 0) throw new InvalidDataException("Empty OPC path segment");
            else segments.Add(segment);
        }
        if (segments.Count == 0) throw new InvalidDataException("Empty OPC target");
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
