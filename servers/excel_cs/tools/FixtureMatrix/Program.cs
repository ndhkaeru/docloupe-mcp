using System.IO.Compression;
using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Schema;
using DocLoupe.Excel.Verify;

if (args.Length is < 1 or > 2)
    throw new ArgumentException("Usage: FixtureMatrix <report.json> [local-corpus-directory]");

var temporary = Path.Combine(Path.GetTempPath(), "docloupe-matrix-" + Guid.NewGuid().ToString("N"));
try
{
    var sources = SyntheticFixtures.Create(Path.Combine(temporary, "synthetic"))
        .Select(path => (Path: path, Corpus: "synthetic")).ToList();
    if (args.Length == 2)
    {
        if (!Directory.Exists(args[1])) throw new DirectoryNotFoundException(args[1]);
        var local = Directory.GetFiles(args[1]).Where(path =>
            Path.GetExtension(path) is ".xlsx" or ".xlsm" &&
            !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal))
            .OrderBy(path => path, StringComparer.Ordinal).ToArray();
        if (local.Length != 8) throw new InvalidDataException($"Expected 8 local fixtures, found {local.Length}");
        sources.AddRange(local.Select(path => (Path: path, Corpus: "local")));
    }
    var rows = sources.Select(source => FixtureMatrix.Check(source.Path, source.Corpus, temporary)).ToArray();
    var report = new
    {
        fixtures = rows,
        summary = new
        {
            total = rows.Length,
            passed = rows.Count(row => row.Status == "passed"),
            gap = rows.Count(row => row.Status == "gap"),
            failed = rows.Count(row => row.Status == "failed")
        }
    };
    var output = Path.GetFullPath(args[0]);
    Directory.CreateDirectory(Path.GetDirectoryName(output)!);
    File.WriteAllText(output, JsonSerializer.Serialize(report, new JsonSerializerOptions
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    }));
    Console.WriteLine($"Fixture matrix: {report.summary.passed} passed, {report.summary.gap} gaps, {report.summary.failed} failed; {output}");
    if (report.summary.failed != 0) Environment.ExitCode = 1;
}
finally
{
    if (Directory.Exists(temporary)) Directory.Delete(temporary, true);
}

internal sealed record GateRow(string Status, string[] Issues, string[] Gaps);
internal sealed record FixtureRow(string Corpus, string Fixture, string Status, Dictionary<string, GateRow> Gates, string? Error);

internal static class FixtureMatrix
{
    public static FixtureRow Check(string source, string corpus, string temporary)
    {
        var gates = new Dictionary<string, GateRow>(StringComparer.Ordinal);
        try
        {
            using var store = new PackageStore(source);
            var sheet = store.SheetNames()[0];
            var unusual = corpus == "local" &&
                (Path.GetFileName(source).StartsWith("06", StringComparison.Ordinal) ||
                 Path.GetFileName(source).StartsWith("07", StringComparison.Ordinal));
            SetValueOp[] operations = corpus == "synthetic"
                ? [new(sheet, "B1", "text", "matrix value")]
                : unusual
                    ? [new(sheet, "A1", "text", "new text"), new(sheet, "B1", "number", "9")]
                    : [new(sheet, "B3", "text", "converted"), new(sheet, "A4", "number", "4"),
                       new(sheet, "C4", "inline", "new inline")];
            var written = Path.Combine(temporary, Guid.NewGuid().ToString("N") + Path.GetExtension(source));
            var result = SetValueEngine.Apply(store, operations);
            store.Save(written);
            var expectation = result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind,
                item.Value, item.AllowMissing, item.RequireMissing, item.KeepCache,
                item.ExplicitCache is { } cache ? new FormulaCacheExpectation(cache.Type, cache.Value) : null)).ToArray();
            var spans = result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End,
                edit.Before, edit.After)).ToArray();
            using var archive = ZipFile.OpenRead(source);
            var originals = archive.Entries.Select(entry => entry.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var changed = result.ChangedParts.Where(part => !store.Contains(part) || !originals.Contains(part));
            gates["G1"] = Gate(P2aGates.CheckPackage(written, result.ChangedParts));
            var schema = DetachedValidator.Check(source, written, result.ChangedParts);
            gates["G2"] = Gate(schema.Issues.Select(issue => issue.Code + ": " + issue.Detail),
                schema.Gaps.Select(issue => issue.Code + ": " + issue.Detail));
            gates["G3"] = Gate(P2aMarkupGate.Check(source, written, result.ChangedParts));
            gates["G4"] = Gate(P2aGates.CheckIntent(written, expectation));
            gates["G5"] = Gate(P2aGates.CheckPreservation(source, written, spans, changed)
                .Concat(P2aGates.CheckSemanticPreservation(source, written, expectation, spans))
                .Concat(P2aGates.CheckTouchedCells(source, written, expectation)));
            var status = gates.Values.Any(gate => gate.Status == "failed") ? "failed" :
                gates.Values.Any(gate => gate.Status == "gap") ? "gap" : "passed";
            return new(corpus, Path.GetFileName(source), status, gates, null);
        }
        catch (Exception exception)
        {
            return new(corpus, Path.GetFileName(source), "failed", gates,
                exception.GetType().Name + ": " + exception.Message);
        }
    }

    private static GateRow Gate(IEnumerable<GateIssue> issues) => Gate(
        issues.Select(issue => issue.Code + ": " + issue.Detail), []);

    private static GateRow Gate(IEnumerable<string> issues, IEnumerable<string> gaps)
    {
        var failures = issues.ToArray();
        var missing = gaps.ToArray();
        return new(failures.Length != 0 ? "failed" : missing.Length != 0 ? "gap" : "passed", failures, missing);
    }
}
