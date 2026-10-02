using System.Globalization;
using System.IO.Compression;
using System.Xml;
using DocLoupe.Excel.Model;

namespace DocLoupe.Excel.Verify;

public sealed record SheetSummary(string Name, string Part, string State, string? UsedRange = null);
public sealed record CellSummary(string Address, string Type, string? RawValue, string? Formula, string? Value = null);
public sealed record WorkbookSummary(string Path, IReadOnlyList<SheetSummary> Sheets, IReadOnlyList<CellSummary> FirstSheetCells);
public sealed record VerificationSummary(string Status, IReadOnlyList<MarkupIssue> PackageIssues,
    IReadOnlyList<MarkupIssue> MarkupIssues, IReadOnlyList<string> UnverifiedGates);

public static class WorkbookReader
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypeNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static WorkbookSummary Peek(string path, int maxCells = 24, string? sheetName = null,
        int maxRows = 1048576, int maxColumns = 16384)
    {
        if (maxCells is < 0 or > 2000) throw new ArgumentOutOfRangeException(nameof(maxCells));
        if (maxRows is < 1 or > 1048576) throw new ArgumentOutOfRangeException(nameof(maxRows));
        if (maxColumns is < 1 or > 16384) throw new ArgumentOutOfRangeException(nameof(maxColumns));
        using var archive = ZipFile.OpenRead(path);
        if (archive.Entries.GroupBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase).Any(group => group.Count() > 1))
            throw new InvalidDataException("Duplicate OPC part");
        var workbookPart = LocateWorkbookPart(archive);
        var relationships = ReadWorkbookRelationships(archive, workbookPart);
        var sheets = ReadSheets(archive, workbookPart, relationships)
            .Select(sheet => sheet with { UsedRange = ReadUsedRange(archive, sheet.Part) }).ToArray();
        var selectedSheet = sheetName is null ? sheets.FirstOrDefault() : sheets.SingleOrDefault(sheet => sheet.Name == sheetName)
            ?? throw new KeyNotFoundException($"Sheet not found: {sheetName}");
        var cells = selectedSheet is null ? [] : ReadCells(archive, workbookPart, selectedSheet.Part, maxCells, maxRows, maxColumns);
        return new WorkbookSummary(path, sheets, cells);
    }

    public static VerificationSummary VerifyPartial(string path)
    {
        try { return VerifyPartialCore(path); }
        catch (InvalidDataException exception)
        {
            return new VerificationSummary("failed", [new MarkupIssue("INVALID_PACKAGE", exception.Message)], [],
                ["G1_REMAINING", "G2", "G4", "G5", "G6", "G7"]);
        }
    }

    private static VerificationSummary VerifyPartialCore(string path)
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

    private static string? ReadUsedRange(ZipArchive archive, string part)
    {
        var entry = FindEntry(archive, part) ?? throw new InvalidDataException($"Missing sheet part {part}");
        using var reader = CreateReader(entry);
        reader.MoveToContent();
        if (reader.LocalName != "worksheet" || reader.NamespaceURI != SpreadsheetNamespace)
            throw new InvalidDataException($"Invalid worksheet root: {part}");
        var sheetDataDepth = -1;
        var rowDepth = -1;
        int? rowNumber = null;
        var minRow = int.MaxValue;
        var minColumn = int.MaxValue;
        var maxRow = 0;
        var maxColumn = 0;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.Depth == rowDepth) { rowDepth = -1; rowNumber = null; }
                if (reader.Depth == sheetDataDepth) sheetDataDepth = -1;
                continue;
            }
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != SpreadsheetNamespace) continue;
            if (reader.LocalName == "sheetData" && reader.Depth == 1)
            {
                sheetDataDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                continue;
            }
            if (sheetDataDepth < 0) continue;
            if (reader.LocalName == "row" && reader.Depth == sheetDataDepth + 1)
            {
                rowDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                var reference = reader.GetAttribute("r");
                rowNumber = null;
                if (reference is not null)
                {
                    if (!int.TryParse(reference, NumberStyles.None, CultureInfo.InvariantCulture, out var parsedRow)
                        || parsedRow is < 1 or > 1048576)
                        throw new InvalidDataException($"Invalid row reference: {reference}");
                    rowNumber = parsedRow;
                }
                continue;
            }
            if (reader.LocalName != "c" || reader.Depth != rowDepth + 1 || rowDepth < 0) continue;
            var address = reader.GetAttribute("r") ?? throw new InvalidDataException("Cell has no address");
            CellAddress coordinate;
            try { coordinate = CellAddress.Parse(address); }
            catch (FormatException exception) { throw new InvalidDataException($"Invalid cell address: {address}", exception); }
            if (rowNumber is not null && coordinate.Row != rowNumber)
                throw new InvalidDataException($"Cell {address} does not belong to row {rowNumber}");
            minRow = Math.Min(minRow, coordinate.Row);
            minColumn = Math.Min(minColumn, coordinate.Column);
            maxRow = Math.Max(maxRow, coordinate.Row);
            maxColumn = Math.Max(maxColumn, coordinate.Column);
        }
        if (maxRow == 0) return null;
        var first = new CellAddress(minRow, minColumn).ToString();
        var last = new CellAddress(maxRow, maxColumn).ToString();
        return first == last ? first : first + ":" + last;
    }

    private static List<CellSummary> ReadCells(ZipArchive archive, string workbookPart, string part,
        int maxCells, int maxRows, int maxColumns)
    {
        var entry = FindEntry(archive, part) ?? throw new InvalidDataException($"Missing sheet part {part}");
        var cells = new List<CellSummary>();
        IReadOnlyList<string>? sharedStrings = null;
        using var reader = CreateReader(entry);
        reader.MoveToContent();
        if (reader.LocalName != "worksheet" || reader.NamespaceURI != SpreadsheetNamespace)
            throw new InvalidDataException("Invalid worksheet root");
        var sheetDataDepth = -1;
        var rowDepth = -1;
        while (cells.Count < maxCells && reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.Depth == rowDepth) rowDepth = -1;
                if (reader.Depth == sheetDataDepth) sheetDataDepth = -1;
                continue;
            }
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != SpreadsheetNamespace) continue;
            if (reader.LocalName == "sheetData" && reader.Depth == 1)
            {
                sheetDataDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                continue;
            }
            if (reader.LocalName == "row" && sheetDataDepth >= 0 && reader.Depth == sheetDataDepth + 1)
            {
                rowDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                continue;
            }
            if (reader.LocalName != "c" || rowDepth < 0 || reader.Depth != rowDepth + 1) continue;
            var address = reader.GetAttribute("r") ?? throw new InvalidDataException("Cell has no address");
            CellAddress coordinate;
            try { coordinate = CellAddress.Parse(address); }
            catch (FormatException exception) { throw new InvalidDataException($"Invalid cell address: {address}", exception); }
            if (coordinate.Row > maxRows || coordinate.Column > maxColumns) continue;
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
            string? resolved = value;
            if (type == "s")
            {
                sharedStrings ??= ReadSharedStrings(archive, workbookPart);
                if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var index)
                    || index < 0 || index >= sharedStrings.Count)
                    throw new InvalidDataException($"Invalid shared string index at {address}: {value}");
                resolved = sharedStrings[index];
            }
            cells.Add(new CellSummary(address, type, value, formula, resolved));
        }
        return cells;
    }

    private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive, string workbookPart)
    {
        var relationshipsPart = FindEntry(archive, RelationshipPart(workbookPart))
            ?? throw new InvalidDataException("Missing workbook relationships");
        string? target = null;
        using (var reader = CreateReader(relationshipsPart))
        {
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship"
                    || reader.NamespaceURI != PackageRelationshipNamespace
                    || reader.GetAttribute("Type") != RelationshipNamespace + "/sharedStrings")
                    continue;
                if (reader.GetAttribute("TargetMode") == "External" || target is not null)
                    throw new InvalidDataException("Invalid shared strings relationship");
                target = reader.GetAttribute("Target") ?? throw new InvalidDataException("Missing shared strings target");
            }
        }
        if (target is null) throw new InvalidDataException("Missing shared strings relationship");
        var part = ResolvePartPath(workbookPart, target);
        var entry = FindEntry(archive, part) ?? throw new InvalidDataException($"Missing shared strings part {part}");
        var values = new List<string>();
        using var stringsReader = CreateReader(entry);
        stringsReader.MoveToContent();
        if (stringsReader.LocalName != "sst" || stringsReader.NamespaceURI != SpreadsheetNamespace)
            throw new InvalidDataException("Invalid shared strings root");
        while (stringsReader.Read())
        {
            if (stringsReader.NodeType != XmlNodeType.Element || stringsReader.LocalName != "si"
                || stringsReader.NamespaceURI != SpreadsheetNamespace)
                continue;
            var text = new System.Text.StringBuilder();
            using (var item = stringsReader.ReadSubtree())
            {
                item.Read();
                while (!item.EOF)
                {
                    if (item.NodeType == XmlNodeType.Element && item.NamespaceURI == SpreadsheetNamespace)
                    {
                        if (item.LocalName is "rPh" or "phoneticPr")
                        {
                            item.Skip();
                            continue;
                        }
                        if (item.LocalName == "t")
                        {
                            text.Append(item.ReadElementContentAsString());
                            continue;
                        }
                    }
                    item.Read();
                }
            }
            values.Add(text.ToString());
        }
        return values;
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
