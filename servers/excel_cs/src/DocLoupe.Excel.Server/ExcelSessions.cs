using System.Collections.Concurrent;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Model;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Schema;
using DocLoupe.Excel.Verify;

namespace DocLoupe.Excel.Server;

public sealed record ReadOnlyVerification(string Path, VerificationSummary Summary, SchemaReport? Schema,
    IReadOnlyList<GateIssue>? AssertionIssues = null);
public sealed record ReadOnlyComparison(ReadOnlyVerification Before, ReadOnlyVerification After,
    PackageComparison? Comparison, SchemaReport? SchemaDelta, IReadOnlyList<GateIssue>? AssertionIssues, string Status);

public sealed class ExcelSessions : IDisposable
{
    private readonly ConcurrentDictionary<string, Session> _sessions = new(StringComparer.Ordinal);
    private static readonly object ServerInfo = new
    {
        version = typeof(ExcelSessions).Assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(attribute => attribute.Key == "DOCLOUPE_SERVER_VERSION")?.Value ?? "development",
        commit = typeof(ExcelSessions).Assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(attribute => attribute.Key == "DOCLOUPE_COMMIT_SHA")?.Value ?? "unknown",
        oracles = new { excel = "unavailable", libreoffice = "unavailable" }
    };

    public object Open(string path)
    {
        var full = Path.GetFullPath(path);
        if (!File.Exists(full)) throw new FileNotFoundException("Workbook not found", full);
        if (Path.GetExtension(full).ToLowerInvariant() is not (".xlsx" or ".xlsm" or ".xltx" or ".xltm"))
            throw new NotSupportedException("Only OOXML workbooks are supported");
        using var store = new PackageStore(full);
        var sheets = store.SheetNames();
        var id = "xs_" + Guid.NewGuid().ToString("N")[..16];
        var session = new Session(id, full, Fingerprint(full));
        if (!_sessions.TryAdd(id, session)) throw new InvalidOperationException("Session collision");
        return new { session = id, revision = 0, path = full, sheets };
    }

    public object CreateFromTemplate(string templatePath, string targetPath)
    {
        var source = Path.GetFullPath(templatePath);
        var destination = Path.GetFullPath(targetPath);
        if (Path.GetExtension(source).ToLowerInvariant() is not (".xlsx" or ".xlsm" or ".xltx" or ".xltm") ||
            !Path.GetExtension(source).Equals(Path.GetExtension(destination), StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Template and output must use the same OOXML workbook format");
        if (!File.Exists(source)) throw new FileNotFoundException("Template not found", source);
        if (!Directory.Exists(Path.GetDirectoryName(destination)))
            throw new DirectoryNotFoundException(Path.GetDirectoryName(destination));
        if (File.Exists(destination)) throw new IOException("Destination already exists");
        var fingerprint = Fingerprint(source);
        using (var template = new PackageStore(source)) template.SheetNames();
        var staging = Path.Combine(Path.GetDirectoryName(destination)!, "." + Path.GetFileName(destination) +
            "." + Guid.NewGuid().ToString("N") + ".staging");
        try
        {
            File.Copy(source, staging);
            if (Fingerprint(source) != fingerprint || Fingerprint(staging) != fingerprint)
                throw new InvalidOperationException("SOURCE_CHANGED_ON_DISK");
            string[] sheets;
            using (var copied = new PackageStore(staging)) sheets = copied.SheetNames().ToArray();
            return PublishCreated(staging, destination, sheets, fingerprint);
        }
        finally { if (File.Exists(staging)) File.Delete(staging); }
    }

    public object CreateNew(string targetPath, string[]? sheets = null, string? activeSheet = null,
        IReadOnlyDictionary<string, string>? coreProperties = null)
    {
        var destination = Path.GetFullPath(targetPath);
        var format = Path.GetExtension(destination).TrimStart('.').ToLowerInvariant();
        if (format is not ("xlsx" or "xlsm" or "xltx" or "xltm"))
            throw new NotSupportedException("Unsupported workbook format");
        if (!Directory.Exists(Path.GetDirectoryName(destination)))
            throw new DirectoryNotFoundException(Path.GetDirectoryName(destination));
        if (File.Exists(destination)) throw new IOException("Destination already exists");
        var staging = Path.Combine(Path.GetDirectoryName(destination)!, "." + Path.GetFileName(destination) +
            "." + Guid.NewGuid().ToString("N") + ".staging");
        try
        {
            var createdSheets = NewWorkbook.Write(staging, sheets, activeSheet, format, coreProperties);
            using (var package = new PackageStore(staging))
                if (!package.SheetNames().SequenceEqual(createdSheets))
                    throw new InvalidDataException("Created worksheets disagree with package relationships");
            var issues = P2aGates.CheckPackage(staging, [], Path.GetExtension(destination));
            var schema = DetachedValidator.CheckPackage(staging);
            if (issues.Count > 0 || schema.Issues.Count > 0 || schema.Gaps.Count > 0)
                throw new InvalidDataException("Created workbook failed OPC or schema validation: " +
                    JsonSerializer.Serialize(new { Package = issues, Schema = schema.Issues, schema.Gaps }));
            return PublishCreated(staging, destination, createdSheets, Fingerprint(staging));
        }
        finally { if (File.Exists(staging)) File.Delete(staging); }
    }

    private object PublishCreated(string staging, string destination, string[] sheets, string fingerprint)
    {
        File.Move(staging, destination);
        string id;
        do { id = "xs_" + Guid.NewGuid().ToString("N")[..16]; }
        while (!_sessions.TryAdd(id, new Session(id, destination, fingerprint)));
        return new { session = id, revision = 0, path = destination, sheets, @new = true,
            default_path = destination };
    }

    public object Peek(string path, string detail = "summary", string? sheet = null, int maxRows = 20, int maxCols = 10)
    {
        if (detail is not ("info" or "summary" or "preview")) throw new ArgumentException("Invalid peek detail");
        if (maxRows is < 1 or > 100 || maxCols is < 1 or > 20 || maxRows * maxCols > 2000)
            throw new ArgumentOutOfRangeException(nameof(maxRows), "Preview is limited to 100 rows, 20 columns and 2000 cells");
        var full = Path.GetFullPath(path);
        if (!File.Exists(full)) throw new FileNotFoundException("Workbook not found", full);
        if (Path.GetExtension(full).ToLowerInvariant() is not (".xlsx" or ".xlsm" or ".xltx" or ".xltm"))
            throw new NotSupportedException("Only OOXML workbooks are supported");
        var workbook = WorkbookReader.Peek(full, detail == "preview" ? maxRows * maxCols : 0, sheet, maxRows, maxCols);
        var sheets = workbook.Sheets.Select((item, index) => new { name = item.Name, index, state = item.State,
            part = item.Part, used_range = item.UsedRange }).ToArray();
        if (detail != "preview")
            return new { sheets, used_range_basis = "explicit_cells", partial = true, unverified = new[] { "features" } };
        var selected = sheet ?? workbook.Sheets.FirstOrDefault()?.Name;
        var preview = selected is null ? Array.Empty<object>() : new object[] { new { sheet = selected, markdown = PreviewMarkdown(workbook.FirstSheetCells, maxRows, maxCols) } };
        return new { sheets, preview, used_range_basis = "explicit_cells", partial = true, unverified = new[] { "features" } };
    }

    private static string PreviewMarkdown(IReadOnlyList<CellSummary> cells, int rows, int columns)
    {
        var lookup = cells.ToDictionary(cell => cell.Address, StringComparer.OrdinalIgnoreCase);
        var text = new System.Text.StringBuilder("| row |");
        for (var column = 1; column <= columns; column++)
            text.Append(' ').Append(new CellAddress(1, column).ToString()[..^1]).Append(" |");
        text.AppendLine().Append("| --- |");
        for (var column = 0; column < columns; column++) text.Append(" --- |");
        for (var row = 1; row <= rows; row++)
        {
            text.AppendLine().Append("| ").Append(row).Append(" |");
            for (var column = 1; column <= columns; column++)
            {
                var address = new CellAddress(row, column).ToString();
                lookup.TryGetValue(address, out var cell);
                var value = cell?.Value ?? (cell?.Formula is null ? "" : "=" + cell.Formula);
                text.Append(' ').Append(value.Replace("\\", "\\\\", StringComparison.Ordinal)
                    .Replace("|", "\\|", StringComparison.Ordinal)
                    .Replace("\r\n", "<br>", StringComparison.Ordinal)
                    .Replace('\n', ' ').Replace('\r', ' ')).Append(" |");
            }
        }
        return text.ToString();
    }

    public ReadOnlyVerification Verify(string afterPath, IReadOnlyList<ValueAssertion>? assertions = null)
    {
        var full = Path.GetFullPath(afterPath);
        if (!File.Exists(full)) throw new FileNotFoundException("Workbook not found", full);
        if (Path.GetExtension(full).ToLowerInvariant() is not (".xlsx" or ".xlsm" or ".xltx" or ".xltm"))
            throw new NotSupportedException("Only OOXML workbooks are supported");
        var summary = WorkbookReader.VerifyPartial(full);
        if (summary.Status == "failed") return new ReadOnlyVerification(full, summary, null);
        var schema = DetachedValidator.CheckPackage(full);
        IReadOnlyList<GateIssue>? assertionIssues = assertions is { Count: > 0 } && schema.Issues.Count == 0
            ? CheckReadOnlyAssertions(full, assertions) : null;
        var status = schema.Issues.Count + (assertionIssues?.Count ?? 0) > 0 ? "failed" : summary.Status;
        return new ReadOnlyVerification(full, summary with { Status = status }, schema, assertionIssues);
    }

    public ReadOnlyComparison Verify(string afterPath, string beforePath, int maxDifferences = 200,
        IReadOnlyList<ValueAssertion>? assertions = null)
    {
        if (maxDifferences is < 1 or > 5000) throw new ArgumentOutOfRangeException(nameof(maxDifferences));
        var before = Verify(beforePath);
        var after = Verify(afterPath);
        if (before.Summary.PackageIssues.Count + before.Summary.MarkupIssues.Count
            + after.Summary.PackageIssues.Count + after.Summary.MarkupIssues.Count > 0)
            return new ReadOnlyComparison(before, after, null, null, null, "failed");
        var comparison = PackageComparator.Compare(before.Path, after.Path, maxDifferences);
        var schemaDelta = DetachedValidator.ComparePackages(before.Path, after.Path);
        IReadOnlyList<GateIssue>? assertionIssues = assertions is { Count: > 0 }
            ? CheckReadOnlyAssertions(after.Path, assertions, before.Path) : null;
        return new ReadOnlyComparison(before, after, comparison, schemaDelta, assertionIssues,
            comparison.HasDifferences || schemaDelta.Issues.Count + (assertionIssues?.Count ?? 0) > 0 ? "failed" : "unverified");
    }

    private static IReadOnlyList<GateIssue> CheckReadOnlyAssertions(string path,
        IReadOnlyList<ValueAssertion> assertions, string? source = null)
    {
        try { return G7Assertions.Check(path, assertions, source); }
        catch (InvalidDataException exception)
        {
            return [new GateIssue("G7", "ASSERT_READ_ERROR", exception.Message)];
        }
    }

    public object Status(string? id = null)
    {
        if (id is null)
            return new { sessions = _sessions.Values.Select(session =>
            {
                var snapshot = session.Snapshot;
                return new { session = session.Id, path = session.Path, revision = snapshot.Revision,
                    dirty = snapshot.Revision != snapshot.SavedRevision };
            }).OrderBy(session => session.session, StringComparer.Ordinal).ToArray(), server = ServerInfo };

        var selected = Get(id);
        var state = selected.Snapshot;
        return new { path = selected.Path, revision = state.Revision, saved_revision = state.SavedRevision,
            dirty = state.Revision != state.SavedRevision, read_only = false,
            source_changed_on_disk = selected.SourceChangedOnDisk(),
            busy = selected.Busy is { } active ? new { operation = active.Operation, since = active.Since } : null,
            ledger = state.Ledger.Select(entry => new { revision = entry.Revision, op_count = entry.OpCount,
                summary = entry.Summary }).ToArray(), server = ServerInfo };
    }

    public object Read(string id, string? sheet, string[] addresses, bool skipEmpty = true, string view = "cells")
    {
        var session = Get(id);
        lock (session.Sync)
        {
            session.CheckSource();
            if (addresses.Length == 0)
            {
                var preview = session.Preview();
                try
                {
                    var workbook = WorkbookReader.Peek(preview, maxCells: 0, sheetName: sheet);
                    var selected = sheet is null ? workbook.Sheets.FirstOrDefault() :
                        workbook.Sheets.Single(item => item.Name == sheet);
                    if (selected is null) throw new InvalidDataException("Workbook has no readable worksheet");
                    if (selected.UsedRange is null)
                        return view switch
                        {
                            "cells" => new { session = id, revision = session.Revision, sheet = selected.Name,
                                view = "cells", cells = Array.Empty<CellRead>() } as object,
                            "values" => new { session = id, revision = session.Revision, sheet = selected.Name,
                                view = "values", rows = Array.Empty<object?[]>() },
                            "markdown" => new { session = id, revision = session.Revision, sheet = selected.Name,
                                view = "markdown", markdown = "" },
                            _ => throw new NotSupportedException("Unsupported read view")
                        };
                    addresses = [selected.Name + "!" + selected.UsedRange];
                }
                finally { if (preview != session.BasePath) File.Delete(preview); }
            }
            if (view is not ("cells" or "values" or "markdown"))
                throw new NotSupportedException("Only cells, values and markdown views are supported");
            if (view != "cells" && addresses.Length != 1)
                throw new NotSupportedException("Values and markdown views require one rectangular range");
            var targets = new List<(string Sheet, string Address)>();
            foreach (var address in addresses)
            {
                var bounds = address.Split(':');
                if (bounds.Length is < 1 or > 2 || bounds.Length == 2 && bounds[1].Contains('!'))
                    throw new FormatException("Invalid cell range");
                var name = CellAddress.SheetName(bounds[0]) ?? sheet
                    ?? throw new ArgumentException("Missing sheet name");
                var first = CellAddress.Parse(bounds[0]);
                var last = bounds.Length == 2 ? CellAddress.Parse(bounds[1]) : first;
                if (last.Row < first.Row || last.Column < first.Column)
                    throw new FormatException("Reversed cell range");
                var count = (long)(last.Row - first.Row + 1) * (last.Column - first.Column + 1);
                if (targets.Count + count > 500) throw new ArgumentException("Read exceeds 500 cells");
                for (var row = first.Row; row <= last.Row; row++)
                    for (var column = first.Column; column <= last.Column; column++)
                        targets.Add((name, new CellAddress(row, column).ToString()));
            }
            var selectedSheet = targets[0].Sheet;
            if (targets.Any(target => target.Sheet != selectedSheet))
                throw new ArgumentException("All cells in a read must belong to one sheet");
            var source = session.Preview();
            try
            {
                var existing = P2aGates.ReadCells(source, selectedSheet, targets.Select(target => target.Address));
                if (view is "values" or "markdown")
                {
                    var indexed = new Dictionary<string, CellRead>(StringComparer.Ordinal);
                    foreach (var cell in existing)
                        if (!indexed.TryAdd(cell.Address, cell))
                            throw new InvalidDataException($"Duplicate cell address: {selectedSheet}!{cell.Address}");
                    var bounds = addresses[0].Split(':');
                    var first = CellAddress.Parse(bounds[0]);
                    var last = bounds.Length == 2 ? CellAddress.Parse(bounds[1]) : first;
                    if (view == "markdown")
                        return new { session = id, revision = session.Revision, sheet = selectedSheet,
                            view = "markdown", markdown = ReadMarkdown(first, last, indexed) };
                    var width = last.Column - first.Column + 1;
                    var rows = targets.Chunk(width).Select(row => row.Select(target =>
                        indexed.TryGetValue(target.Address, out var cell) ? TypedValue(cell) : null).ToArray()).ToArray();
                    return new { session = id, revision = session.Revision, sheet = selectedSheet,
                        view = "values", rows };
                }
                IReadOnlyList<CellRead> cells = existing;
                if (!skipEmpty)
                {
                    var indexed = new Dictionary<string, CellRead>(StringComparer.Ordinal);
                    foreach (var cell in existing)
                        if (!indexed.TryAdd(cell.Address, cell))
                            throw new InvalidDataException($"Duplicate cell address: {selectedSheet}!{cell.Address}");
                    cells = targets.Select(target => indexed.TryGetValue(target.Address, out var cell)
                        ? cell : new CellRead(target.Address, "blank", null, null)).ToArray();
                }
                return new { session = id, revision = session.Revision, sheet = selectedSheet, view = "cells", cells };
            }
            finally { if (source != session.BasePath) File.Delete(source); }
        }
    }

    public object Find(string id, string? sheet, string? target, string pattern, bool isRegex,
        string searchIn = "value", bool caseSensitive = false, string normalize = "nfc", int maxResults = 100,
        ValueAssertion? expectedValue = null)
    {
        if (searchIn is not ("value" or "formula") || normalize is not ("nfc" or "none") ||
            expectedValue is not null && (searchIn != "value" || isRegex || pattern.Length != 0))
            throw new NotSupportedException("Only value and formula search with nfc or none normalization is supported");
        if (maxResults is < 1 or > 100 || expectedValue is null && pattern.Length is < 1 or > 512)
            throw new ArgumentOutOfRangeException(nameof(maxResults), "Search accepts 1..100 results and a 1..512 character pattern");
        var needle = normalize == "nfc" ? pattern.Normalize(System.Text.NormalizationForm.FormC) : pattern;
        var regex = isRegex ? new System.Text.RegularExpressions.Regex(needle,
            System.Text.RegularExpressions.RegexOptions.CultureInvariant |
            (caseSensitive ? System.Text.RegularExpressions.RegexOptions.None : System.Text.RegularExpressions.RegexOptions.IgnoreCase),
            TimeSpan.FromMilliseconds(100)) : null;
        var session = Get(id);
        lock (session.Sync)
        {
            session.CheckSource();
            var source = session.Preview();
            try
            {
                var summary = target is null ? WorkbookReader.Peek(source, maxCells: 0) : null;
                var ranges = new List<(string Sheet, string[] Addresses)>();
                if (target is not null)
                    ranges.Add(FindTargets(sheet, target));
                else
                {
                    if (sheet is not null && !summary!.Sheets.Any(item => item.Name == sheet))
                        throw new KeyNotFoundException("Sheet not found: " + sheet);
                    var remaining = 500;
                    foreach (var item in summary!.Sheets.Where(item => sheet is null || item.Name == sheet))
                    {
                        if (item.UsedRange is null) continue;
                        var range = FindTargets(item.Name, item.UsedRange, remaining);
                        remaining -= range.Addresses.Length;
                        ranges.Add(range);
                    }
                }
                var matches = new List<object>();
                var truncated = false;
                var scanned = 0;
                foreach (var (selectedSheet, addresses) in ranges)
                {
                    var cells = P2aGates.ReadCells(source, selectedSheet, addresses)
                        .ToDictionary(cell => cell.Address, StringComparer.Ordinal);
                    foreach (var address in addresses)
                    {
                        scanned++;
                        cells.TryGetValue(address, out var cell);
                        var found = false;
                        if (expectedValue is not null)
                            found = FindValueMatches(cell, expectedValue, caseSensitive, normalize);
                        else if (cell is not null)
                        {
                            var raw = searchIn == "formula" ? cell.Formula : cell.Value;
                            if (raw is null) continue;
                            var haystack = normalize == "nfc" ? raw.Normalize(System.Text.NormalizationForm.FormC) : raw;
                            found = regex is not null ? regex.IsMatch(haystack) :
                                haystack.Contains(needle, caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
                        }
                        if (!found) continue;
                        if (matches.Count == maxResults) { truncated = true; break; }
                        matches.Add(new { addr = selectedSheet + "!" + address,
                            value = cell is null ? null : TypedValue(cell), formula = cell?.Formula });
                    }
                    if (truncated) break;
                }
                return new { session = id, revision = session.Revision, partial = true,
                    total_scanned = scanned, matches, truncated };
            }
            finally { if (source != session.BasePath) File.Delete(source); }
        }
    }

    private static bool FindValueMatches(CellRead? cell, ValueAssertion expected, bool caseSensitive, string normalize)
    {
        if (expected.Kind != "text") return G7Assertions.ValueMatches(cell, expected);
        var text = cell?.Kind == "formula" && cell.CacheType == "str" ? cell.CacheRawValue :
            cell?.Kind is "text" or "inline" ? cell.Value : null;
        if (text is null) return false;
        var actual = normalize == "nfc" ? text.Normalize(System.Text.NormalizationForm.FormC) : text;
        var wanted = normalize == "nfc" ? expected.Value!.Normalize(System.Text.NormalizationForm.FormC) : expected.Value!;
        return string.Equals(actual, wanted, caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
    }

    private static (string Sheet, string[] Addresses) FindTargets(string? sheet, string target, int maxCells = 500)
    {
        var bounds = target.Split(':');
        if (bounds.Length is < 1 or > 2 || bounds.Length == 2 && bounds[1].Contains('!'))
            throw new FormatException("Invalid search range");
        var selectedSheet = CellAddress.SheetName(bounds[0]) ?? sheet ?? throw new ArgumentException("Missing search sheet");
        var first = CellAddress.Parse(bounds[0]);
        var last = bounds.Length == 2 ? CellAddress.Parse(bounds[1]) : first;
        if (last.Row < first.Row || last.Column < first.Column) throw new FormatException("Reversed search range");
        var count = (long)(last.Row - first.Row + 1) * (last.Column - first.Column + 1);
        if (count > maxCells) throw new ArgumentException("Search exceeds 500 cells; supply a smaller scope.target");
        var addresses = new List<string>((int)count);
        for (var row = first.Row; row <= last.Row; row++)
            for (var column = first.Column; column <= last.Column; column++)
                addresses.Add(new CellAddress(row, column).ToString());
        return (selectedSheet, addresses.ToArray());
    }

    private static string ReadMarkdown(CellAddress first, CellAddress last,
        IReadOnlyDictionary<string, CellRead> cells)
    {
        var text = new System.Text.StringBuilder("| row |");
        for (var column = first.Column; column <= last.Column; column++)
            text.Append(' ').Append(new CellAddress(1, column).ToString()[..^1]).Append(" |");
        text.AppendLine().Append("| --- |");
        for (var column = first.Column; column <= last.Column; column++) text.Append(" --- |");
        for (var row = first.Row; row <= last.Row; row++)
        {
            text.AppendLine().Append("| ").Append(row).Append(" |");
            for (var column = first.Column; column <= last.Column; column++)
            {
                cells.TryGetValue(new CellAddress(row, column).ToString(), out var cell);
                var value = cell?.Value ?? "";
                if (cell?.Formula is { } formula) value += " ƒ =" + formula;
                if (cell is not null && IsRich(cell)) value = "† " + value;
                text.Append(' ').Append(value.Replace("\\", "\\\\", StringComparison.Ordinal)
                    .Replace("|", "\\|", StringComparison.Ordinal)
                    .Replace("<", "&lt;", StringComparison.Ordinal)
                    .Replace(">", "&gt;", StringComparison.Ordinal)
                    .Replace("\r\n", "<br>", StringComparison.Ordinal)
                    .Replace("\n", "<br>", StringComparison.Ordinal)
                    .Replace("\r", "<br>", StringComparison.Ordinal)).Append(" |");
            }
        }
        return text.ToString();
    }

    private static bool IsRich(CellRead cell) =>
        System.Text.RegularExpressions.Regex.IsMatch(cell.SharedMarkup ?? cell.CellMarkup ?? "",
            @"<([A-Za-z_][\w.-]*:)?r(?=[\s>/])", System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    private static object? TypedValue(CellRead cell)
    {
        if (cell.Kind == "blank" || cell.Kind == "formula" && cell.Value is null) return null;
        var kind = cell.Kind == "formula" ? cell.CacheType switch
        {
            null or "" or "n" => "number", "str" => "text", "b" => "boolean", "e" => "error",
            _ => throw new NotSupportedException("Unsupported formula cache type in values view")
        } : cell.Kind;
        return kind switch
        {
            "text" or "inline" => cell.Value,
            "boolean" when cell.Value is "true" or "false" => cell.Value == "true",
            "error" when cell.Value is not null => new { error = cell.Value },
            "number" when cell.Value is not null => ParseNumber(cell.Value),
            _ => throw new InvalidDataException("Invalid cell value in values view: " + cell.Address)
        };
    }

    private static JsonElement ParseNumber(string value)
    {
        using var document = JsonDocument.Parse(value);
        if (document.RootElement.ValueKind != JsonValueKind.Number)
            throw new InvalidDataException("Invalid number in values view");
        return document.RootElement.Clone();
    }

    public object Apply(string id, int baseRevision, SetValueOp[] operations, bool dryRun = false,
        int maxDiffItems = 200, string returnMode = "diff+readback")
    {
        var session = Get(id);
        lock (session.Sync)
        {
            session.CheckSource();
            if (operations.Length is < 1 or > 500) throw new ArgumentOutOfRangeException(nameof(operations));
            if (maxDiffItems is < 1 or > 500) throw new ArgumentOutOfRangeException(nameof(maxDiffItems));
            if (returnMode is not ("diff" or "diff+readback"))
                throw new NotSupportedException("Unsupported apply return mode");
            if (baseRevision != session.Revision) throw new InvalidOperationException("REVISION_CONFLICT");
            if (operations.Any(operation => operation.Expect is not null))
            {
                var basePath = session.Preview();
                try
                {
                    var preconditions = operations.Where(operation => operation.Expect is not null).ToArray();
                    var observed = new Dictionary<string, CellRead?>(StringComparer.Ordinal);
                    var missingSheets = new HashSet<string>(StringComparer.Ordinal);
                    foreach (var group in preconditions.GroupBy(operation => operation.Sheet, StringComparer.Ordinal))
                    {
                        try
                        {
                            var cells = P2aGates.ReadCells(basePath, group.Key,
                                group.Select(operation => operation.Address).Distinct(StringComparer.Ordinal))
                                .ToDictionary(cell => cell.Address, StringComparer.Ordinal);
                            foreach (var operation in group)
                                observed[operation.Sheet + "!" + operation.Address] = cells.GetValueOrDefault(operation.Address);
                        }
                        catch (InvalidOperationException) { missingSheets.Add(group.Key); }
                    }
                    for (var index = 0; index < operations.Length; index++)
                    {
                        var operation = operations[index];
                        if (operation.Expect is not { } expected) continue;
                        var target = operation.Sheet + "!" + operation.Address;
                        observed.TryGetValue(target, out var actual);
                        var assertionMatches = G7Assertions.Matches(actual, new ValueAssertion(
                            operation.Sheet, operation.Address, expected.CheckValue,
                            expected.Kind, expected.Value, expected.Formula));
                        var isEmpty = actual is null || actual.Kind == "blank" && actual.Formula is null;
                        if (assertionMatches && !missingSheets.Contains(operation.Sheet) &&
                            (expected.Empty is null || expected.Empty == isEmpty) &&
                            (expected.Text is null || actual is { Kind: "text" or "inline" } &&
                                string.Equals(actual.Value, expected.Text, StringComparison.Ordinal))) continue;
                        throw new PreconditionFailedException(operation.SourceIndex < 0 ? index : operation.SourceIndex,
                            target, expected, actual);
                    }
                }
                finally { if (basePath != session.BasePath) File.Delete(basePath); }
            }
            using var candidate = new PackageStore(session.BasePath);
            var next = Coalesce(session.Operations.Concat(operations));
            var result = SetValueEngine.Apply(candidate, next);
            var readback = Readback(candidate, operations);
            IReadOnlyDictionary<string, CellRead?> previous;
            var priorPath = session.Preview();
            try { previous = Readback(priorPath, operations); }
            finally { if (priorPath != session.BasePath) File.Delete(priorPath); }
            var (diff, diffSummary) = DescribeDiff(previous, readback, maxDiffItems);
            var results = operations.Select((operation, index) =>
            {
                var result = new Dictionary<string, object?>
                {
                    ["index"] = operation.SourceIndex < 0 ? index : operation.SourceIndex,
                    ["op"] = operation.Operation,
                    ["status"] = dryRun ? "planned" : "applied",
                    ["resolved"] = operation.Sheet + "!" + operation.Address
                };
                if (operation.Label is not null) result["label"] = operation.Label;
                return result;
            }).ToArray();
            if (dryRun)
            {
                if (returnMode == "diff")
                    return new { session = id, dry_run = true, revision = session.Revision,
                        revision_before = baseRevision, revision_after = baseRevision,
                        intent = result.Intent, changed_parts = result.ChangedParts, results,
                        diff, diff_summary = diffSummary };
                return new { session = id, dry_run = true, revision = session.Revision,
                    revision_before = baseRevision, revision_after = baseRevision,
                    intent = result.Intent, changed_parts = result.ChangedParts, readback, results,
                    diff, diff_summary = diffSummary };
            }
            session.Operations.AddRange(operations);
            session.RevisionLengths.Add(operations.Length);
            session.Revision++;
            session.Ledger.Add(new LedgerEntry(session.Revision, operations.Length,
                string.Join(", ", operations.Select(operation => $"{operation.Operation} {operation.Sheet}!{operation.Address}"))));
            session.Publish();
            if (returnMode == "diff")
                return new { session = id, dry_run = false, revision = session.Revision,
                    revision_before = baseRevision, revision_after = session.Revision,
                    intent = result.Intent, changed_parts = result.ChangedParts, results,
                    diff, diff_summary = diffSummary };
            return new { session = id, dry_run = false, revision = session.Revision,
                revision_before = baseRevision, revision_after = session.Revision,
                intent = result.Intent, changed_parts = result.ChangedParts, readback, results,
                diff, diff_summary = diffSummary };
        }
    }

    private static IReadOnlyDictionary<string, CellRead?> Readback(PackageStore candidate, SetValueOp[] operations)
    {
        var staging = Path.Combine(Path.GetTempPath(), "docloupe-readback-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            candidate.Save(staging);
            return Readback(staging, operations);
        }
        finally { if (File.Exists(staging)) File.Delete(staging); }
    }

    private static IReadOnlyDictionary<string, CellRead?> Readback(string path, SetValueOp[] operations)
    {
        var readback = new Dictionary<string, CellRead?>(StringComparer.Ordinal);
        foreach (var group in operations.GroupBy(operation => operation.Sheet, StringComparer.Ordinal))
        {
            var addresses = group.Select(operation => operation.Address).Distinct(StringComparer.Ordinal).ToArray();
            var cells = P2aGates.ReadCells(path, group.Key, addresses)
                .ToDictionary(cell => cell.Address, StringComparer.Ordinal);
            foreach (var address in addresses)
                readback[group.Key + "!" + address] = cells.GetValueOrDefault(address);
        }
        return readback;
    }

    private static (object[] Diff, object Summary) DescribeDiff(
        IReadOnlyDictionary<string, CellRead?> previous, IReadOnlyDictionary<string, CellRead?> readback,
        int maxDiffItems)
    {
        var changes = new List<(string Path, string Facet, object? Before, object? After,
            string BeforeJson, string AfterJson)>();
        foreach (var (target, after) in readback)
        {
            previous.TryGetValue(target, out var before);
            Add("present", before is not null, after is not null);
            Add("type", before?.Kind, after?.Kind);
            Add("value", before?.Value, after?.Value);
            Add("formula", before?.Formula, after?.Formula);
            Add("cache", before?.Formula is null ? null : new { type = before.CacheType, raw = before.CacheRawValue },
                after?.Formula is null ? null : new { type = after.CacheType, raw = after.CacheRawValue });

            void Add(string facet, object? oldValue, object? newValue)
            {
                var oldJson = JsonSerializer.Serialize(oldValue);
                var newJson = JsonSerializer.Serialize(newValue);
                if (oldJson != newJson)
                    changes.Add((target + "." + facet, facet, oldValue, newValue, oldJson, newJson));
            }
        }
        var diff = changes.Take(maxDiffItems).Select(change => new
        {
            id = "d_" + Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(
                change.Path + change.BeforeJson + change.AfterJson))).ToLowerInvariant()[..10],
            path = change.Path,
            before = change.Before,
            after = change.After
        }).ToArray();
        var summary = new
        {
            facets_changed = changes.Count,
            cells_touched = readback.Count,
            by_category = changes.GroupBy(change => change.Facet, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal),
            truncated = changes.Count > maxDiffItems,
            partial = true
        };
        return (diff, summary);
    }

    public UndoResult Undo(string id, int baseRevision, int toRevision)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            session.CheckSource();
            if (baseRevision != session.Revision) throw new InvalidOperationException("REVISION_CONFLICT");
            if (toRevision < 0 || toRevision > session.Revision)
                throw new ArgumentOutOfRangeException(nameof(toRevision));
            var discarded = Enumerable.Range(toRevision + 1, session.Revision - toRevision).ToArray();
            var keptOperations = session.RevisionLengths.Take(toRevision).Sum();
            if (keptOperations > 0)
            {
                using var candidate = new PackageStore(session.BasePath);
                SetValueEngine.Apply(candidate, Coalesce(session.Operations.Take(keptOperations)));
            }
            session.Operations.RemoveRange(keptOperations, session.Operations.Count - keptOperations);
            session.RevisionLengths.RemoveRange(toRevision, session.RevisionLengths.Count - toRevision);
            session.Ledger.RemoveRange(toRevision, session.Ledger.Count - toRevision);
            session.Revision = toRevision;
            session.Publish();
            return new UndoResult(session.Revision, discarded);
        }
    }

    public object Save(string id, string? outputPath, IReadOnlyList<ValueAssertion>? assertions = null, string mode = "copy")
    {
        if (mode is not ("copy" or "save_as" or "overwrite")) throw new NotSupportedException("Unsupported save mode");
        var session = Get(id);
        lock (session.Sync)
        {
            if (session.Revision == 0 && session.SavedRevision == 0)
                throw new InvalidOperationException("No pending edits");
            session.CheckSource();
            var destination = mode == "overwrite" && outputPath is null ? session.Path :
                Path.GetFullPath(outputPath ?? throw new ArgumentException("Save path is required"));
            var samePath = string.Equals(destination, session.Path,
                OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
            if (mode == "overwrite" ? !samePath : samePath)
                throw new ArgumentException("Save mode and destination path disagree");
            if (!Path.GetExtension(destination).Equals(Path.GetExtension(session.Path), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Output format must match source format");
            if (!Directory.Exists(Path.GetDirectoryName(destination))) throw new DirectoryNotFoundException(Path.GetDirectoryName(destination));
            if (mode != "overwrite" && File.Exists(destination)) throw new IOException("Destination already exists");
            var staging = Path.Combine(Path.GetDirectoryName(destination)!, "." + Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".staging");
            var backup = mode == "overwrite" ? Path.Combine(Path.GetDirectoryName(destination)!,
                "." + Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".bak") : null;
            session.SetBusy("save");
            try
            {
                using var store = new PackageStore(session.BasePath);
                if (AdvancedPartGate.CheckSignedSource(session.BasePath) is { Count: > 0 } signed)
                    throw new SaveBlockedException(signed);
                var result = session.Operations.Count == 0 ? new ApplyResult([], [], []) :
                    SetValueEngine.Apply(store, Coalesce(session.Operations));
                store.Save(staging);
                var reports = new List<GateIssue>();
                reports.AddRange(P2aGates.CheckPackage(staging, result.ChangedParts, Path.GetExtension(session.BasePath)));
                var schema = DetachedValidator.Check(session.BasePath, staging, result.ChangedParts);
                reports.AddRange(schema.Issues.Select(issue => new GateIssue("G2", issue.Code, issue.Detail)));
                reports.AddRange(P2aMarkupGate.Check(session.BasePath, staging, result.ChangedParts));
                reports.AddRange(P2aGates.CheckIntent(staging, result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value, item.AllowMissing, item.RequireMissing, item.KeepCache,
                    item.ExplicitCache is { } cache ? new FormulaCacheExpectation(cache.Type, cache.Value) : null))));
                var addedOrRemoved = result.ChangedParts.Where(part => !store.Contains(part) || !PartExists(session.BasePath, part));
                reports.AddRange(P2aGates.CheckPreservation(session.BasePath, staging,
                    result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), addedOrRemoved));
                reports.AddRange(P2aGates.CheckTouchedCells(session.BasePath, staging,
                    result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value, item.AllowMissing, item.RequireMissing, item.KeepCache,
                    item.ExplicitCache is { } cache ? new FormulaCacheExpectation(cache.Type, cache.Value) : null))));
                reports.AddRange(P2aGates.CheckSemanticPreservation(session.BasePath, staging,
                    result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value, item.AllowMissing, item.RequireMissing, item.KeepCache,
                    item.ExplicitCache is { } cache ? new FormulaCacheExpectation(cache.Type, cache.Value) : null)),
                    result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After))));
                reports.AddRange(AdvancedPartGate.Check(session.BasePath, staging,
                    result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value, item.AllowMissing, item.RequireMissing, item.KeepCache,
                    item.ExplicitCache is { } cache ? new FormulaCacheExpectation(cache.Type, cache.Value) : null))));
                if (reports.Count == 0) reports.AddRange(G7Assertions.Check(staging, assertions ?? [], session.BasePath));
                if (reports.Count > 0) throw new SaveBlockedException(reports);
                if (schema.Gaps.Count > 0) throw new SaveBlockedException(schema.Gaps.Select(issue => new GateIssue("G2", issue.Code, issue.Detail)).ToArray());
                var readback = result.Intent.GroupBy(item => item.Sheet).ToDictionary(group => group.Key,
                    group => P2aGates.ReadCells(staging, group.Key, group.Select(item => item.Address)));
                session.CheckSource();
                var writtenFingerprint = Fingerprint(staging);
                var priorFingerprint = session.CurrentFingerprint;
                string? baselineSnapshot = null;
                try
                {
                    if (mode != "copy") baselineSnapshot = session.PrepareBaseline();
                    session.CheckSource();
                    if (mode == "overwrite") File.Replace(staging, destination, backup);
                    else File.Move(staging, destination);
                    if (mode != "copy")
                    {
                        session.CommitFollow(destination, writtenFingerprint, baselineSnapshot);
                        baselineSnapshot = null;
                    }
                }
                finally { Session.DiscardPreparedBaseline(baselineSnapshot); }
                return new { session = id, revision = session.Revision, revision_saved = mode == "copy" ?
                        session.SavedRevision : session.Revision, path = destination, status = "verified",
                    backup = backup is null ? null : new { path = backup, sha256 = priorFingerprint },
                    gates = new[] { "G1", "G2", "G3", "G4", "G5", "G6", "G7" },
                    assertions = (assertions ?? []).Select((item, index) => new { index, status = "verified",
                        expected = item }).ToArray(), readback };
            }
            finally
            {
                session.ClearBusy();
                if (File.Exists(staging)) File.Delete(staging);
            }
        }
    }

    public object Close(string id, bool discardUnsaved)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            if (session.Revision != session.SavedRevision && !discardUnsaved) throw new InvalidOperationException("UNSAVED_CHANGES");
            session.Cleanup();
            _sessions.TryRemove(id, out _);
            return new { closed = true, discarded_revisions = session.Revision };
        }
    }

    private static SetValueOp[] Coalesce(IEnumerable<SetValueOp> operations) => operations
        .GroupBy(operation => (operation.Sheet, CellAddress.Parse(operation.Address)))
        .Select(group => group.Last()).ToArray();

    private static string Fingerprint(string path)
    {
        using var source = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(source));
    }

    private Session Get(string id) => _sessions.TryGetValue(id, out var session) ? session : throw new KeyNotFoundException("Unknown session");

    private static bool PartExists(string file, string part)
    {
        using var zip = System.IO.Compression.ZipFile.OpenRead(file);
        return zip.Entries.Any(entry => entry.FullName.Equals(part, StringComparison.OrdinalIgnoreCase));
    }

    public void Dispose()
    {
        foreach (var session in _sessions.Values) session.Cleanup();
        _sessions.Clear();
    }

    private sealed class Session(string id, string path, string fingerprint)
    {
        public string Id { get; } = id;
        public string Path { get; private set; } = path;
        public string BasePath { get; private set; } = path;
        public string BaseFingerprint { get; } = fingerprint;
        public string CurrentFingerprint { get; private set; } = fingerprint;
        public int SavedRevision { get; private set; }
        private string? _snapshotDirectory;
        public object Sync { get; } = new();
        public List<SetValueOp> Operations { get; } = [];
        public List<int> RevisionLengths { get; } = [];
        public List<LedgerEntry> Ledger { get; } = [];
        public int Revision { get; set; }
        private SessionSnapshot _snapshot = new(0, 0, []);
        private BusyOperation? _busy;
        public SessionSnapshot Snapshot => Volatile.Read(ref _snapshot);
        public BusyOperation? Busy => Volatile.Read(ref _busy);

        public void Publish() => Volatile.Write(ref _snapshot,
            new SessionSnapshot(Revision, SavedRevision, Ledger.TakeLast(20).ToArray()));

        public void SetBusy(string operation) => Volatile.Write(ref _busy,
            new BusyOperation(operation, DateTimeOffset.UtcNow.ToString("O")));

        public void ClearBusy() => Volatile.Write(ref _busy, null);

        public bool SourceChangedOnDisk()
        {
            try { return ExcelSessions.Fingerprint(BasePath) != BaseFingerprint ||
                BasePath != Path && ExcelSessions.Fingerprint(Path) != CurrentFingerprint; }
            catch (IOException) { return true; }
            catch (UnauthorizedAccessException) { return true; }
        }

        public void CheckSource()
        {
            if (SourceChangedOnDisk()) throw new InvalidOperationException("SOURCE_CHANGED_ON_DISK");
        }

        public string? PrepareBaseline()
        {
            if (BasePath != Path) return null;
            var directory = Directory.CreateTempSubdirectory("docloupe-excel-session-").FullName;
            var snapshot = System.IO.Path.Combine(directory, System.IO.Path.GetFileName(Path));
            try
            {
                File.Copy(Path, snapshot);
                if (ExcelSessions.Fingerprint(snapshot) != BaseFingerprint)
                    throw new InvalidOperationException("SOURCE_CHANGED_ON_DISK");
                return snapshot;
            }
            catch
            {
                Directory.Delete(directory, true);
                throw;
            }
        }

        public void CommitFollow(string destination, string fingerprint, string? baselineSnapshot)
        {
            if (baselineSnapshot is not null)
            {
                _snapshotDirectory = System.IO.Path.GetDirectoryName(baselineSnapshot);
                BasePath = baselineSnapshot;
            }
            Path = destination;
            CurrentFingerprint = fingerprint;
            SavedRevision = Revision;
            Publish();
        }

        public static void DiscardPreparedBaseline(string? snapshot)
        {
            if (snapshot is not null) Directory.Delete(System.IO.Path.GetDirectoryName(snapshot)!, true);
        }

        public void Cleanup()
        {
            if (_snapshotDirectory is { } directory)
            {
                Directory.Delete(directory, true);
                _snapshotDirectory = null;
            }
        }

        public string Preview()
        {
            if (Revision == 0) return BasePath;
            using var candidate = new PackageStore(BasePath);
            SetValueEngine.Apply(candidate, Coalesce(Operations));
            var temporary = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "docloupe-p2a-" + Guid.NewGuid().ToString("N") + ".xlsx");
            candidate.Save(temporary);
            return temporary;
        }
    }
}

public sealed record UndoResult(int Revision, IReadOnlyList<int> Discarded);

public sealed record LedgerEntry(int Revision, int OpCount, string Summary);

public sealed record SessionSnapshot(int Revision, int SavedRevision, IReadOnlyList<LedgerEntry> Ledger);

public sealed record BusyOperation(string Operation, string Since);

public sealed class PreconditionFailedException(int index, string target, CellPrecondition expected, CellRead? actual)
    : Exception("PRECONDITION_FAILED: " + target)
{
    public int Index { get; } = index;
    public string Target { get; } = target;
    public CellPrecondition Expected { get; } = expected;
    public CellRead? Actual { get; } = actual;
}

public sealed class SaveBlockedException(IReadOnlyList<GateIssue> issues) : Exception("SAVE_BLOCKED: " + JsonSerializer.Serialize(issues))
{
    public IReadOnlyList<GateIssue> Issues { get; } = issues;
}
