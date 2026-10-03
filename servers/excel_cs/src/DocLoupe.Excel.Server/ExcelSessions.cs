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
                    dirty = snapshot.Revision != 0 };
            }).OrderBy(session => session.session, StringComparer.Ordinal).ToArray(), server = ServerInfo };

        var selected = Get(id);
        var state = selected.Snapshot;
        return new { path = selected.Path, revision = state.Revision, saved_revision = 0,
            dirty = state.Revision != 0, read_only = false,
            source_changed_on_disk = selected.SourceChangedOnDisk(),
            busy = selected.Busy is { } active ? new { operation = active.Operation, since = active.Since } : null,
            ledger = state.Ledger.Select(entry => new { revision = entry.Revision, op_count = entry.OpCount,
                summary = entry.Summary }).ToArray(), server = ServerInfo };
    }

    public object Read(string id, string? sheet, string[] addresses, bool skipEmpty = true)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            session.CheckSource();
            if (addresses.Length == 0) throw new ArgumentException("At least one cell address is required");
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
            finally { if (source != session.Path) File.Delete(source); }
        }
    }

    public object Apply(string id, int baseRevision, SetValueOp[] operations)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            session.CheckSource();
            if (operations.Length is < 1 or > 500) throw new ArgumentOutOfRangeException(nameof(operations));
            if (baseRevision != session.Revision) throw new InvalidOperationException("REVISION_CONFLICT");
            if (operations.Any(operation => operation.Expect is not null))
            {
                var basePath = session.Preview();
                try
                {
                    for (var index = 0; index < operations.Length; index++)
                    {
                        var operation = operations[index];
                        if (operation.Expect is not { } expected) continue;
                        var assertion = new ValueAssertion(operation.Sheet, operation.Address, expected.CheckValue,
                            expected.Kind, expected.Value, expected.Formula);
                        if (G7Assertions.Check(basePath, [assertion]).Count == 0) continue;
                        CellRead? actual;
                        try { actual = P2aGates.ReadCells(basePath, operation.Sheet, [operation.Address]).SingleOrDefault(); }
                        catch (InvalidOperationException) { actual = null; }
                        throw new PreconditionFailedException(operation.SourceIndex < 0 ? index : operation.SourceIndex,
                            operation.Sheet + "!" + operation.Address, expected, actual);
                    }
                }
                finally { if (basePath != session.Path) File.Delete(basePath); }
            }
            using var candidate = new PackageStore(session.Path);
            var next = Coalesce(session.Operations.Concat(operations));
            var result = SetValueEngine.Apply(candidate, next);
            session.Operations.AddRange(operations);
            session.RevisionLengths.Add(operations.Length);
            session.Revision++;
            session.Ledger.Add(new LedgerEntry(session.Revision, operations.Length,
                string.Join(", ", operations.Select(operation => $"{operation.Operation} {operation.Sheet}!{operation.Address}"))));
            session.Publish();
            return new { session = id, revision = session.Revision, intent = result.Intent };
        }
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
                using var candidate = new PackageStore(session.Path);
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

    public object Save(string id, string outputPath, IReadOnlyList<ValueAssertion>? assertions = null)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            if (session.Revision == 0) throw new InvalidOperationException("No pending edits");
            session.CheckSource();
            var destination = Path.GetFullPath(outputPath);
            if (destination == session.Path) throw new NotSupportedException("P2a requires a distinct output path; overwrite is not yet supported");
            if (!Path.GetExtension(destination).Equals(Path.GetExtension(session.Path), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Output format must match source format");
            if (!Directory.Exists(Path.GetDirectoryName(destination))) throw new DirectoryNotFoundException(Path.GetDirectoryName(destination));
            if (File.Exists(destination)) throw new IOException("Destination already exists");
            var staging = Path.Combine(Path.GetDirectoryName(destination)!, "." + Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".staging");
            session.SetBusy("save");
            try
            {
                using var store = new PackageStore(session.Path);
                if (AdvancedPartGate.CheckSignedSource(session.Path) is { Count: > 0 } signed)
                    throw new SaveBlockedException(signed);
                var result = SetValueEngine.Apply(store, Coalesce(session.Operations));
                store.Save(staging);
                var reports = new List<GateIssue>();
                reports.AddRange(P2aGates.CheckPackage(staging, result.ChangedParts, Path.GetExtension(session.Path)));
                var schema = DetachedValidator.Check(session.Path, staging, result.ChangedParts);
                reports.AddRange(schema.Issues.Select(issue => new GateIssue("G2", issue.Code, issue.Detail)));
                reports.AddRange(P2aMarkupGate.Check(session.Path, staging, result.ChangedParts));
                reports.AddRange(P2aGates.CheckIntent(staging, result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value, item.AllowMissing, item.RequireMissing, item.KeepCache,
                    item.ExplicitCache is { } cache ? new FormulaCacheExpectation(cache.Type, cache.Value) : null))));
                var addedOrRemoved = result.ChangedParts.Where(part => !store.Contains(part) || !PartExists(session.Path, part));
                reports.AddRange(P2aGates.CheckPreservation(session.Path, staging,
                    result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), addedOrRemoved));
                reports.AddRange(P2aGates.CheckTouchedCells(session.Path, staging,
                    result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value, item.AllowMissing, item.RequireMissing, item.KeepCache,
                    item.ExplicitCache is { } cache ? new FormulaCacheExpectation(cache.Type, cache.Value) : null))));
                reports.AddRange(AdvancedPartGate.Check(session.Path, staging,
                    result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value, item.AllowMissing, item.RequireMissing, item.KeepCache,
                    item.ExplicitCache is { } cache ? new FormulaCacheExpectation(cache.Type, cache.Value) : null))));
                if (reports.Count == 0) reports.AddRange(G7Assertions.Check(staging, assertions ?? [], session.Path));
                if (reports.Count > 0) throw new SaveBlockedException(reports);
                if (schema.Gaps.Count > 0) throw new SaveBlockedException(schema.Gaps.Select(issue => new GateIssue("G2", issue.Code, issue.Detail)).ToArray());
                var readback = result.Intent.GroupBy(item => item.Sheet).ToDictionary(group => group.Key,
                    group => P2aGates.ReadCells(staging, group.Key, group.Select(item => item.Address)));
                var response = new { session = id, revision = session.Revision, path = destination, status = "verified",
                    gates = new[] { "G1", "G2", "G3", "G4", "G5", "G6", "G7" },
                    assertions = (assertions ?? []).Select((item, index) => new { index, status = "verified",
                        expected = item }).ToArray(), readback };
                File.Move(staging, destination);
                return response;
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
            if (session.Revision > 0 && !discardUnsaved) throw new InvalidOperationException("UNSAVED_CHANGES");
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
        _sessions.Clear();
    }

    private sealed class Session(string id, string path, string fingerprint)
    {
        public string Id { get; } = id;
        public string Path { get; } = path;
        public string Fingerprint { get; } = fingerprint;
        public object Sync { get; } = new();
        public List<SetValueOp> Operations { get; } = [];
        public List<int> RevisionLengths { get; } = [];
        public List<LedgerEntry> Ledger { get; } = [];
        public int Revision { get; set; }
        private SessionSnapshot _snapshot = new(0, []);
        private BusyOperation? _busy;
        public SessionSnapshot Snapshot => Volatile.Read(ref _snapshot);
        public BusyOperation? Busy => Volatile.Read(ref _busy);

        public void Publish() => Volatile.Write(ref _snapshot,
            new SessionSnapshot(Revision, Ledger.TakeLast(20).ToArray()));

        public void SetBusy(string operation) => Volatile.Write(ref _busy,
            new BusyOperation(operation, DateTimeOffset.UtcNow.ToString("O")));

        public void ClearBusy() => Volatile.Write(ref _busy, null);

        public bool SourceChangedOnDisk()
        {
            try { return ExcelSessions.Fingerprint(Path) != Fingerprint; }
            catch (IOException) { return true; }
            catch (UnauthorizedAccessException) { return true; }
        }

        public void CheckSource()
        {
            if (ExcelSessions.Fingerprint(Path) != Fingerprint)
                throw new InvalidOperationException("SOURCE_CHANGED_ON_DISK");
        }

        public string Preview()
        {
            if (Revision == 0) return Path;
            using var candidate = new PackageStore(Path);
            SetValueEngine.Apply(candidate, Coalesce(Operations));
            var temporary = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "docloupe-p2a-" + Guid.NewGuid().ToString("N") + ".xlsx");
            candidate.Save(temporary);
            return temporary;
        }
    }
}

public sealed record UndoResult(int Revision, IReadOnlyList<int> Discarded);

public sealed record LedgerEntry(int Revision, int OpCount, string Summary);

public sealed record SessionSnapshot(int Revision, IReadOnlyList<LedgerEntry> Ledger);

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
