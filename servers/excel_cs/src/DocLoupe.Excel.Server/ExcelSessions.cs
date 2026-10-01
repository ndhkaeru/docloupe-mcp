using System.Collections.Concurrent;
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

    public object Open(string path)
    {
        var full = Path.GetFullPath(path);
        if (!File.Exists(full)) throw new FileNotFoundException("Workbook not found", full);
        if (Path.GetExtension(full).ToLowerInvariant() is not (".xlsx" or ".xlsm" or ".xltx" or ".xltm"))
            throw new NotSupportedException("Only OOXML workbooks are supported");
        var store = new PackageStore(full);
        var id = "xs_" + Guid.NewGuid().ToString("N")[..16];
        var session = new Session(id, full, Fingerprint(full), store);
        if (!_sessions.TryAdd(id, session)) throw new InvalidOperationException("Session collision");
        return new { session = id, revision = 0, path = full, sheets = store.SheetNames() };
    }

    public object Read(string id, string sheet, string[] addresses)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            session.CheckSource();
            var source = session.Preview();
            try { return new { session = id, revision = session.Revision, sheet, view = "cells", cells = P2aGates.ReadCells(source, sheet, addresses) }; }
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
            using var candidate = new PackageStore(session.Path);
            var next = Coalesce(session.Operations.Concat(operations));
            var result = SetValueEngine.Apply(candidate, next);
            session.Operations.AddRange(operations);
            session.Revision++;
            return new { session = id, revision = session.Revision, intent = result.Intent };
        }
    }

    public object Save(string id, string outputPath)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            if (session.Revision == 0) throw new InvalidOperationException("No pending edits");
            session.CheckSource();
            if (session.Store.Parts.Any(part => part.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
                throw new SaveBlockedException([new GateIssue("G6", "SIGNED_PACKAGE_UNSUPPORTED", "P2a cannot safely update signed workbooks")]);
            var destination = Path.GetFullPath(outputPath);
            if (destination == session.Path) throw new NotSupportedException("P2a requires a distinct output path; overwrite is not yet supported");
            if (!Path.GetExtension(destination).Equals(Path.GetExtension(session.Path), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Output format must match source format");
            if (!Directory.Exists(Path.GetDirectoryName(destination))) throw new DirectoryNotFoundException(Path.GetDirectoryName(destination));
            if (File.Exists(destination)) throw new IOException("Destination already exists");
            var staging = Path.Combine(Path.GetDirectoryName(destination)!, "." + Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".staging");
            try
            {
                using var store = new PackageStore(session.Path);
                var result = SetValueEngine.Apply(store, Coalesce(session.Operations));
                store.Save(staging);
                var reports = new List<GateIssue>();
                reports.AddRange(P2aGates.CheckPackage(staging, result.ChangedParts, Path.GetExtension(session.Path)));
                var schema = DetachedValidator.Check(session.Path, staging, result.ChangedParts);
                reports.AddRange(schema.Issues.Select(issue => new GateIssue("G2", issue.Code, issue.Detail)));
                reports.AddRange(P2aMarkupGate.Check(session.Path, staging, result.ChangedParts));
                reports.AddRange(P2aGates.CheckIntent(staging, result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value))));
                var addedOrRemoved = result.ChangedParts.Where(part => !store.Contains(part) || !PartExists(session.Path, part));
                reports.AddRange(P2aGates.CheckPreservation(session.Path, staging,
                    result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), addedOrRemoved));
                reports.AddRange(P2aGates.CheckTouchedCells(session.Path, staging,
                    result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value))));
                if (reports.Count > 0) throw new SaveBlockedException(reports);
                if (schema.Gaps.Count > 0) throw new SaveBlockedException(schema.Gaps.Select(issue => new GateIssue("G2", issue.Code, issue.Detail)).ToArray());
                var readback = result.Intent.GroupBy(item => item.Sheet).ToDictionary(group => group.Key,
                    group => P2aGates.ReadCells(staging, group.Key, group.Select(item => item.Address)));
                File.Move(staging, destination);
                return new { session = id, revision = session.Revision, path = destination, status = "verified",
                    gates = new[] { "G1", "G2", "G3", "G4", "G5" }, readback };
            }
            finally { if (File.Exists(staging)) File.Delete(staging); }
        }
    }

    public object Close(string id, bool discardUnsaved)
    {
        var session = Get(id);
        lock (session.Sync)
        {
            if (session.Revision > 0 && !discardUnsaved) throw new InvalidOperationException("UNSAVED_CHANGES");
            _sessions.TryRemove(id, out _);
            session.Store.Dispose();
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
        foreach (var session in _sessions.Values) session.Store.Dispose();
        _sessions.Clear();
    }

    private sealed class Session(string id, string path, string fingerprint, PackageStore store)
    {
        public string Id { get; } = id;
        public string Path { get; } = path;
        public string Fingerprint { get; } = fingerprint;
        public PackageStore Store { get; } = store;
        public object Sync { get; } = new();
        public List<SetValueOp> Operations { get; } = [];
        public int Revision { get; set; }

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

public sealed class SaveBlockedException(IReadOnlyList<GateIssue> issues) : Exception("SAVE_BLOCKED: " + JsonSerializer.Serialize(issues))
{
    public IReadOnlyList<GateIssue> Issues { get; } = issues;
}
