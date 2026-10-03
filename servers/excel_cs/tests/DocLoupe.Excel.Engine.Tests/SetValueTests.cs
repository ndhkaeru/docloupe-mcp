using System.IO.Compression;
using System.Text;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Model;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Schema;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SetValueTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("nested-workbook")]
    public void BatchRoundTripAndGates(string variant)
    {
        using var fixture = new Fixture();
        var source = fixture.Source(variant);
        var output = Path.Combine(fixture.Directory, "written.xlsx");
        var operations = new[]
        {
            new SetValueOp("Sheet1", "A1", "text", "sửa & mới", "replace"),
            new SetValueOp("Sheet1", "B1", "text", "converted"),
            new SetValueOp("Sheet1", "C1", "number", "37"),
            new SetValueOp("Sheet1", "D3", "inline", " inline "),
            new SetValueOp("Sheet1", "A2", "text", "new row"),
            new SetValueOp("Sheet1", "B2", "number", "2.5"),
            new SetValueOp("Sheet1", "C3", "formula", "=1+2")
        };
        using var store = new PackageStore(source);
        var result = SetValueEngine.Apply(store, operations);
        store.Save(output);
        Assert.Empty(P2aGates.CheckPackage(output, result.ChangedParts));
        Assert.Empty(P2aGates.CheckIntent(output, result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value))));
        Assert.Empty(P2aGates.CheckPreservation(source, output,
            result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)),
            result.ChangedParts.Where(part => !FileContainsPart(source, part))));
        Assert.Empty(DetachedValidator.Check(source, output, result.ChangedParts).Issues);
    }

    [Theory]
    [InlineData("A1", "number", "12")]
    [InlineData("A1", "formula", "=1+3")]
    [InlineData("A1", "inline", "inline text")]
    [InlineData("B1", "text", "shared text")]
    [InlineData("B1", "formula", "=1+3")]
    [InlineData("B1", "inline", "inline text")]
    [InlineData("C1", "text", "shared text")]
    [InlineData("C1", "number", "12")]
    [InlineData("C1", "inline", "inline text")]
    [InlineData("D3", "text", "shared text")]
    [InlineData("D3", "number", "12")]
    [InlineData("D3", "formula", "=1+3")]
    public void ConvertsBetweenAllFourCellKinds(string address, string kind, string value)
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        var output = Path.Combine(fixture.Directory, "converted.xlsx");
        using var store = new PackageStore(source);
        var result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", address, kind, value, "replace")]);
        store.Save(output);
        var expected = new CellExpectation("Sheet1", address, kind, value);
        Assert.Empty(P2aGates.CheckPackage(output, result.ChangedParts));
        Assert.Empty(P2aGates.CheckIntent(output, [expected]));
        Assert.Empty(P2aGates.CheckTouchedCells(source, output, [expected]));
        Assert.Empty(P2aGates.CheckPreservation(source, output, result.Edits.Select(edit =>
            new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), []));
        Assert.Empty(P2aMarkupGate.Check(source, output, result.ChangedParts));
        Assert.Empty(DetachedValidator.Check(source, output, result.ChangedParts).Issues);
    }

    [Fact]
    public void RejectsEditingMergedNonOrigin()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
            AddBeforeClose(archive, "xl/worksheets/sheet1.xml", "</worksheet>",
                "<mergeCells count=\"1\"><mergeCell ref=\"A1:C1\"/></mergeCells>");
        using var store = new PackageStore(source);
        var error = Assert.Throws<InvalidDataException>(() => SetValueEngine.Apply(store,
            [new SetValueOp("Sheet1", "B1", "number", "3")]));
        Assert.Contains("MERGED_NON_ORIGIN", error.Message);
        Assert.Empty(store.ChangedParts);
    }

    [Fact]
    public void SharedStringReadbackExcludesPhoneticGuideText()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        Assert.Equal("hello", Assert.Single(P2aGates.ReadCells(source, "Sheet1", ["A1"])).Value);
        Assert.Empty(P2aGates.CheckIntent(source, [new CellExpectation("Sheet1", "A1", "text", "hello")]));
    }

    [Fact]
    public void RichOrPhoneticCellRequiresExplicitReplacement()
    {
        using var fixture = new Fixture();
        using var store = new PackageStore(fixture.Source("default"));
        Assert.Throws<InvalidDataException>(() => SetValueEngine.Apply(store,
            [new SetValueOp("Sheet1", "A1", "text", "new")]));
        Assert.Empty(store.ChangedParts);
        var result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "A1", "text", "new", "replace")]);
        Assert.Contains(result.Intent, item => item.Address == "A1");
    }

    [Theory]
    [InlineData("new-shared-strings", "xl/sharedStrings.xml")]
    [InlineData("nested-workbook", "xl/nested/sharedStrings.xml")]
    public void CreatesSharedStringsPartAndRelationshipWhenAbsent(string variant, string sharedPart)
    {
        using var fixture = new Fixture();
        var source = fixture.Source(variant);
        var output = Path.Combine(fixture.Directory, "added-sst.xlsx");
        using var store = new PackageStore(source);
        var result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "text", "new string")]);
        store.Save(output);
        Assert.Contains(sharedPart, result.ChangedParts);
        Assert.Empty(P2aGates.CheckPackage(output, result.ChangedParts));
        Assert.Empty(P2aGates.CheckIntent(output, [new CellExpectation("Sheet1", "B1", "text", "new string")]));
        Assert.Empty(P2aGates.CheckPreservation(source, output, result.Edits.Select(edit =>
            new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), [sharedPart]));
        Assert.Empty(DetachedValidator.Check(source, output, result.ChangedParts).Issues);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FormulaOverwriteRemovesCalcChainButOtherEditsKeepIt(bool explicitEndTags)
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
        {
            var relationship = "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/>";
            var overrideTag = "<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>";
            if (explicitEndTags)
            {
                relationship = relationship.Replace("/>", "></Relationship>", StringComparison.Ordinal);
                overrideTag = overrideTag.Replace("/>", "></Override>", StringComparison.Ordinal);
            }
            AddBeforeClose(archive, "xl/_rels/workbook.xml.rels", "</Relationships>", relationship);
            AddBeforeClose(archive, "[Content_Types].xml", "</Types>", overrideTag);
            using var stream = archive.CreateEntry("xl/calcChain.xml").Open();
            stream.Write(Encoding.UTF8.GetBytes("<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"C1\" i=\"1\"/></calcChain>"));
        }
        foreach (var (address, kind, value, expectedRemoval) in new[]
            { ("B1", "number", "9", false), ("C1", "number", "9", true), ("C1", "formula", "1+3", false) })
        {
            var output = Path.Combine(fixture.Directory, address + "-" + kind + ".xlsx");
            using var store = new PackageStore(source);
            var result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", address, kind, value)]);
            store.Save(output);
            using var zip = ZipFile.OpenRead(output);
            Assert.Equal(!expectedRemoval, zip.GetEntry("xl/calcChain.xml") is not null);
            Assert.Empty(P2aGates.CheckPackage(output, result.ChangedParts));
            var spans = result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)).ToArray();
            Assert.Empty(P2aGates.CheckPreservation(source, output, spans, expectedRemoval ? ["xl/calcChain.xml"] : []));
            Assert.Empty(P2aGates.CheckSemanticPreservation(source, output,
                [new CellExpectation("Sheet1", address, kind, value)], spans));
        }
    }

    [Fact]
    public void NewSharedStringsAndCalcChainDeletionDoNotHideOtherBytes()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("new-shared-strings");
        using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
        {
            AddBeforeClose(archive, "xl/_rels/workbook.xml.rels", "</Relationships>", "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/>");
            AddBeforeClose(archive, "[Content_Types].xml", "</Types>", "<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>");
            using var stream = archive.CreateEntry("xl/calcChain.xml").Open();
            stream.Write(Encoding.UTF8.GetBytes("<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"C1\" i=\"1\"/></calcChain>"));
        }
        var output = Path.Combine(fixture.Directory, "combined.xlsx");
        using var store = new PackageStore(source);
        var result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "text", "new"), new SetValueOp("Sheet1", "C1", "number", "8")]);
        store.Save(output);
        var expected = result.Intent.Select(item => new CellExpectation(item.Sheet, item.Address, item.Kind, item.Value));
        var spans = result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End,
            edit.Before, edit.After)).ToArray();
        Assert.Empty(P2aGates.CheckPackage(output, result.ChangedParts));
        Assert.Empty(P2aGates.CheckPreservation(source, output, spans, ["xl/calcChain.xml", "xl/sharedStrings.xml"]));
        Assert.Empty(P2aGates.CheckSemanticPreservation(source, output, expected, spans));
    }

    private static void AddBeforeClose(ZipArchive archive, string part, string closing, string addition)
    {
        var entry = archive.GetEntry(part)!;
        byte[] original;
        using (var input = entry.Open()) { using var memory = new MemoryStream(); input.CopyTo(memory); original = memory.ToArray(); }
        entry.Delete();
        using var output = archive.CreateEntry(part).Open();
        output.Write(Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(original).Replace(closing, addition + closing, StringComparison.Ordinal)));
    }

    [Fact]
    public void LostEditFailsG4EvenWhenBytesElsewhereArePreserved()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("prefixed-x");
        var output = Path.Combine(fixture.Directory, "lost.xlsx");
        File.Copy(source, output);
        var issues = P2aGates.CheckIntent(output, [new CellExpectation("Sheet1", "A1", "text", "not there")]);
        Assert.Contains(issues, issue => issue.Code == "INTENT_MISMATCH");
    }

    [Fact]
    public void UndeclaredPartChangeFailsG5()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        var output = Path.Combine(fixture.Directory, "changed.xlsx");
        File.Copy(source, output);
        using (var archive = ZipFile.Open(output, ZipArchiveMode.Update))
        {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            byte[] bytes;
            using (var stream = entry.Open()) { using var memory = new MemoryStream(); stream.CopyTo(memory); bytes = memory.ToArray(); }
            entry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var writer = replacement.Open();
            writer.Write(Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace("old", "lost", StringComparison.Ordinal)));
        }
        Assert.Contains(P2aGates.CheckPreservation(source, output, [], []), issue => issue.Code == "UNDECLARED_PART_CHANGE");
    }

    [Theory]
    [InlineData("A0")]
    [InlineData("XFE1")]
    [InlineData("A1:B2")]
    [InlineData("A01")]
    public void InvalidAddressFails(string address)
    {
        Assert.ThrowsAny<Exception>(() => CellAddress.Parse(address));
    }

    private static bool FileContainsPart(string file, string part)
    {
        using var archive = ZipFile.OpenRead(file);
        return archive.Entries.Any(entry => entry.FullName.Equals(part, StringComparison.OrdinalIgnoreCase));
    }

    private sealed class Fixture : IDisposable
    {
        public string Directory { get; } = Path.Combine(Path.GetTempPath(), "docloupe-p2a-" + Guid.NewGuid().ToString("N"));
        public string Source(string name)
        {
            SyntheticFixtures.Create(Directory);
            return Path.Combine(Directory, name + ".xlsx");
        }
        public void Dispose() => System.IO.Directory.Delete(Directory, true);
    }
}
