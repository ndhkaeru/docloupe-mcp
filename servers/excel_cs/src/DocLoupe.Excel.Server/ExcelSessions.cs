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

    public object Read(string id, string? sheet, string[] addresses)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            session.CheckSource();
            if (addresses.Length == 0) throw new ArgumentException("At least one cell address is required");
            var targets = addresses.Select(address => (Sheet: CellAddress.SheetName(address) ?? sheet
                ?? throw new ArgumentException("Missing sheet name"), Address: CellAddress.Parse(address).ToString())).ToArray();
            var selectedSheet = targets[0].Sheet;
            if (targets.Any(target => target.Sheet != selectedSheet))
                throw new ArgumentException("All cells in a read must belong to one sheet");
            var source = session.Preview();
            try { return new { session = id, revision = session.Revision, sheet = selectedSheet, view = "cells",
                cells = P2aGates.ReadCells(source, selectedSheet, targets.Select(target => target.Address)) }; }
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
                        var assertion = new ValueAssertion(operation.Sheet, operation.Address, true, expected.Kind, expected.Value, null);
                        if (G7Assertions.Check(basePath, [assertion]).Count == 0) continue;
                        CellRead? actual;
                        try { actual = P2aGates.ReadCells(basePath, operation.Sheet, [operation.Address]).SingleOrDefault(); }
                        catch (InvalidOperationException) { actual = null; }
                        throw new PreconditionFailedException(index, operation.Sheet + "!" + operation.Address, expected, actual);
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

public sealed class PreconditionFailedException(int index, string target, ExpectedValue expected, CellRead? actual)
    : Exception("PRECONDITION_FAILED: " + target)
{
    public int Index { get; } = index;
    public string Target { get; } = target;
    public ExpectedValue Expected { get; } = expected;
    public CellRead? Actual { get; } = actual;
}

public sealed class SaveBlockedException(IReadOnlyList<GateIssue> issues) : Exception("SAVE_BLOCKED: " + JsonSerializer.Serialize(issues))
{
    public IReadOnlyList<GateIssue> Issues { get; } = issues;
}
