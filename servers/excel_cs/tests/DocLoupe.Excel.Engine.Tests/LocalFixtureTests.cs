using System.IO.Compression;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Schema;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class LocalFixtureTests
{
    [Fact]
    public void LocalSourcesAcceptMultiCellSetValueMatrix()
    {
        var directory = Environment.GetEnvironmentVariable("DOCLOUPE_P2A_LOCAL_FIXTURES");
        if (string.IsNullOrWhiteSpace(directory)) return;
        var sources = Directory.GetFiles(directory, "*.*")
            .Where(path => Path.GetExtension(path) is ".xlsx" or ".xlsm")
            .Where(path => !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal))
            .OrderBy(path => path).ToArray();
        Assert.Equal(8, sources.Length);
        foreach (var source in sources)
        {
            var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + Path.GetExtension(source));
            try
            {
                using var store = new PackageStore(source);
                var sheet = store.SheetNames()[0];
                var unusual = Path.GetFileName(source).StartsWith("06", StringComparison.Ordinal) ||
                    Path.GetFileName(source).StartsWith("07", StringComparison.Ordinal);
                SetValueOp[] ops = unusual
                    ? [new(sheet, "A1", "text", "new text"), new(sheet, "B1", "number", "9")]
                    : [new(sheet, "B3", "text", "converted"), new(sheet, "A4", "number", "4"), new(sheet, "C4", "inline", "new inline")];
                var result = SetValueEngine.Apply(store, ops);
                store.Save(output);
                Assert.Empty(P2aGates.CheckPackage(output, result.ChangedParts));
                Assert.Empty(P2aGates.CheckIntent(output, result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value))));
                Assert.Empty(P2aGates.CheckPreservation(source, output, result.Edits.Select(edit =>
                    new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), AddedOrRemoved(source, store, result.ChangedParts)));
                Assert.Empty(P2aMarkupGate.Check(source, output, result.ChangedParts));
                Assert.Empty(DetachedValidator.Check(source, output, result.ChangedParts).Issues);
            }
            finally { if (File.Exists(output)) File.Delete(output); }
        }
    }

    [Fact]
    public void LocalSourcesAcceptClearValuesWithoutCreatingAbsentCells()
    {
        var directory = Environment.GetEnvironmentVariable("DOCLOUPE_P2A_LOCAL_FIXTURES");
        if (string.IsNullOrWhiteSpace(directory)) return;
        var sources = Directory.GetFiles(directory, "*.*")
            .Where(path => Path.GetExtension(path) is ".xlsx" or ".xlsm")
            .Where(path => !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal))
            .OrderBy(path => path).ToArray();
        Assert.Equal(8, sources.Length);
        foreach (var source in sources)
        {
            var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + Path.GetExtension(source));
            try
            {
                using var store = new PackageStore(source);
                var sheet = store.SheetNames()[0];
                var unusual = Path.GetFileName(source).StartsWith("06", StringComparison.Ordinal) ||
                    Path.GetFileName(source).StartsWith("07", StringComparison.Ordinal);
                var address = unusual ? "A1" : "B3";
                var absent = "XFD1048576";
                var result = SetValueEngine.Apply(store, [new SetValueOp(sheet, address, "blank", null, Operation: "clear"),
                    new SetValueOp(sheet, absent, "blank", null, Operation: "clear")]);
                store.Save(output);
                var expected = result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind,
                    item.Value, item.AllowMissing)).ToArray();
                Assert.Empty(P2aGates.CheckPackage(output, result.ChangedParts));
                Assert.Empty(P2aGates.CheckIntent(output, expected));
                Assert.Empty(P2aGates.ReadCells(output, sheet, [absent]));
                Assert.Empty(P2aGates.CheckTouchedCells(source, output, expected));
                Assert.Empty(P2aGates.CheckPreservation(source, output, result.Edits.Select(edit =>
                    new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)),
                    AddedOrRemoved(source, store, result.ChangedParts)));
                Assert.Empty(P2aMarkupGate.Check(source, output, result.ChangedParts));
                Assert.Empty(DetachedValidator.Check(source, output, result.ChangedParts).Issues);
            }
            finally { if (File.Exists(output)) File.Delete(output); }
        }
    }

    private static IEnumerable<string> AddedOrRemoved(string source, PackageStore store, IEnumerable<string> changedParts)
    {
        using var archive = ZipFile.OpenRead(source);
        var originalParts = archive.Entries.Select(entry => entry.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        return changedParts.Where(part => !store.Contains(part) || !originalParts.Contains(part)).ToArray();
    }

    [Fact]
    public void SignedLocalOriginalBlocksSaveWithoutOutput()
    {
        var directory = Environment.GetEnvironmentVariable("DOCLOUPE_P2A_LOCAL_FIXTURES");
        if (string.IsNullOrWhiteSpace(directory)) return;
        var source = Path.Combine(directory, "05-advanced-package-source.xlsm");
        using var sessions = new ExcelSessions();
        var opened = sessions.Open(source);
        var id = (string)opened.GetType().GetProperty("session")!.GetValue(opened)!;
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsm");
        sessions.Apply(id, 0, [new SetValueOp("Scores", "B3", "number", "19")]);
        var blocked = Assert.Throws<SaveBlockedException>(() => sessions.Save(id, output));
        Assert.Contains(blocked.Issues, issue => issue.Code == "SIGNED_PACKAGE_UNSUPPORTED");
        Assert.False(File.Exists(output));
        sessions.Close(id, true);
    }

    [Fact]
    public void LocalSourcesAreNeverUsedByCiUnlessExplicitlySelected()
    {
        var directory = Environment.GetEnvironmentVariable("DOCLOUPE_P2A_LOCAL_FIXTURES");
        if (string.IsNullOrWhiteSpace(directory)) return;
        var sourceFiles = Directory.GetFiles(directory, "*.*")
            .Where(path => Path.GetExtension(path) is ".xlsx" or ".xlsm")
            .Where(path => !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal))
            .OrderBy(path => path).ToArray();
        Assert.Equal(8, sourceFiles.Length);
        foreach (var source in sourceFiles)
        {
            var temporary = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + Path.GetExtension(source));
            try
            {
                using var package = new PackageStore(source);
                var sheet = package.SheetNames()[0];
                var address = Path.GetFileName(source).StartsWith("06", StringComparison.Ordinal) ||
                    Path.GetFileName(source).StartsWith("07", StringComparison.Ordinal) ? "A1" : "B3";
                var operation = new SetValueOp(sheet, address, "number", "19");
                var result = SetValueEngine.Apply(package, [operation]);
                package.Save(temporary);
                Assert.Empty(P2aGates.CheckPackage(temporary, result.ChangedParts));
                Assert.Empty(P2aGates.CheckIntent(temporary, [new CellExpectation(sheet, address, "number", "19")]));
                Assert.Empty(P2aGates.CheckPreservation(source, temporary, result.Edits.Select(edit =>
                    new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), AddedOrRemoved(source, package, result.ChangedParts)));
                Assert.Empty(P2aMarkupGate.Check(source, temporary, result.ChangedParts));
                Assert.Empty(DetachedValidator.Check(source, temporary, result.ChangedParts).Issues);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
