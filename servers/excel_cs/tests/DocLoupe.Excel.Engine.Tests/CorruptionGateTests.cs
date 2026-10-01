using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Schema;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class CorruptionGateTests
{
    [Fact]
    public void G1RejectsDanglingRelationshipReference()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace("<sheetData", "<sheetData r:id=\"rId999\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", StringComparison.Ordinal));
        Assert.Contains(P2aGates.CheckPackage(broken, ["xl/worksheets/sheet1.xml"]), issue => issue.Gate == "G1" && issue.Code == "DANGLING_REFERENCE");
    }

    [Fact]
    public void G2RejectsIncorrectCellChildOrder()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => Regex.Replace(xml, @"<f>1\+1</f>\s*<v>2</v>", "<v>2</v><f>1+1</f>"));
        Assert.Contains(DetachedValidator.Check(fixture.Source, broken, ["xl/worksheets/sheet1.xml"]).Issues, issue => issue.Code == "NEW_SCHEMA_ERROR");
    }

    [Fact]
    public void G2ReportsMaskedParentWhenBaselineCellErrorExists()
    {
        using var fixture = new Fixture();
        var invalid = fixture.Corrupt("xl/worksheets/sheet1.xml", xml =>
            Regex.Replace(xml, @"<c r=""B1"" t=""n"">\s*<v>42</v>", "<c r=\"B1\" t=\"n\"><v>42</v><f>3</f>"));
        var written = Path.Combine(fixture.Directory, "masked.xlsx");
        using (var store = new PackageStore(invalid))
        {
            SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "7")]);
            store.Save(written);
        }
        Assert.Contains(DetachedValidator.Check(invalid, written, ["xl/worksheets/sheet1.xml"]).Gaps,
            issue => issue.Code == "G2_MASKED_BY_BASELINE_ERROR");
    }

    [Fact]
    public void G3RejectsUndeclaredMcPrefix()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace("mc:Ignorable=\"x14ac\"", "mc:Ignorable=\"x14ac absent\"", StringComparison.Ordinal));
        Assert.Contains(P2aMarkupGate.Check(fixture.Source, broken, ["xl/worksheets/sheet1.xml"]), issue => issue.Gate == "G3" && issue.Code == "UNDECLARED_MC_PREFIX");
    }

    [Fact]
    public void G4RejectsSilentLostEdit()
    {
        using var fixture = new Fixture();
        Assert.Contains(P2aGates.CheckIntent(fixture.Source, [new CellExpectation("Sheet1", "B1", "number", "999")]),
            issue => issue.Gate == "G4" && issue.Code == "INTENT_MISMATCH");
    }

    [Fact]
    public void G5RejectsUnrequestedStyleInsideEditedCell()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace("<c r=\"B1\" t=\"n\">", "<c r=\"B1\" t=\"n\" s=\"5\">", StringComparison.Ordinal));
        Assert.Contains(P2aGates.CheckTouchedCells(fixture.Source, broken, [new CellExpectation("Sheet1", "B1", "number", "42")]),
            issue => issue.Gate == "G5" && issue.Code == "CELL_ATTRIBUTE_CHANGED");
    }

    [Fact]
    public void G5RejectsUndeclaredBytesOutsideDeclaredEdit()
    {
        using var fixture = new Fixture();
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        using (var store = new PackageStore(fixture.Source))
        {
            var result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(edited);
            var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace("old", "damaged", StringComparison.Ordinal), edited);
            Assert.Contains(P2aGates.CheckPreservation(fixture.Source, broken,
                result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), []),
                issue => issue.Gate == "G5" && issue.Code.StartsWith("UNDECLARED", StringComparison.Ordinal));
        }
    }

    private sealed class Fixture : IDisposable
    {
        public string Directory { get; } = Path.Combine(Path.GetTempPath(), "docloupe-corrupt-" + Guid.NewGuid().ToString("N"));
        public string Source { get; }
        public Fixture()
        {
            SyntheticFixtures.Create(Directory);
            Source = Path.Combine(Directory, "default.xlsx");
        }
        public string Corrupt(string part, Func<string, string> transform, string? original = null)
        {
            original ??= Source;
            var result = Path.Combine(Directory, Guid.NewGuid().ToString("N") + ".xlsx");
            using var archive = ZipFile.OpenRead(original);
            using var destination = new ZipArchive(File.Create(result), ZipArchiveMode.Create);
            foreach (var entry in archive.Entries)
            {
                var output = destination.CreateEntry(entry.FullName);
                using var input = entry.Open();
                using var stream = output.Open();
                if (entry.FullName == part)
                {
                    using var reader = new StreamReader(input, Encoding.UTF8);
                    var old = reader.ReadToEnd();
                    var modified = transform(old);
                    Assert.NotEqual(old, modified);
                    stream.Write(Encoding.UTF8.GetBytes(modified));
                }
                else input.CopyTo(stream);
            }
            return result;
        }
        public void Dispose() => System.IO.Directory.Delete(Directory, true);
    }
}
