using System.IO.Compression;
using System.Text;
using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ClearTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void ClearValuesAcrossSyntheticFixturesPreservesAbsentCells(string variant)
    {
        using var fixture = new Fixture();
        var source = fixture.Source(variant);
        using var sessions = new ExcelSessions();
        var id = Open(sessions, source);
        var request = Request("""{"op":"clear","target":"Sheet1!A1:D3","what":["values"],"remove_cells":false}""");
        sessions.Apply(id, 0, request.NormalizeMany(null));
        var output = fixture.Output("cleared");
        var report = sessions.Save(id, output);
        Assert.Equal("verified", report.GetType().GetProperty("status")!.GetValue(report));
        var cells = P2aGates.ReadCells(output, "Sheet1", ["A1", "B1", "C1", "D3", "A2", "D2"]);
        Assert.Equal(4, cells.Count);
        Assert.All(cells, cell => Assert.Equal("blank", cell.Kind));
        Assert.Empty(P2aGates.CheckIntent(output, [new CellExpectation("Sheet1", "D2", "blank", null, true)]));
        Assert.Empty(P2aGates.CheckTouchedCells(source, output, [new CellExpectation("Sheet1", "D2", "blank", null, true)]));
        sessions.Close(id, true);
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void RemovingCellsAcrossSyntheticFixturesKeepsOtherParts(string variant)
    {
        using var fixture = new Fixture();
        var source = fixture.Source(variant);
        using var sessions = new ExcelSessions();
        var id = Open(sessions, source);
        var request = Request("""{"op":"clear","target":"Sheet1!A1:D3","what":["values"],"remove_cells":true}""");
        var applied = sessions.Apply(id, 0, request.NormalizeMany(null));
        var intent = (IReadOnlyList<ExpectedCell>)applied.GetType().GetProperty("intent")!.GetValue(applied)!;
        Assert.All(intent, item => Assert.True(item.RequireMissing));
        var output = fixture.Output("removed");
        sessions.Save(id, output);
        Assert.Empty(P2aGates.ReadCells(output, "Sheet1", ["A1", "B1", "C1", "D3", "A2"]));
        sessions.Close(id, true);
    }

    [Fact]
    public void ClearingStyledCellKeepsStyleAndClearingAbsentCellChangesNoPart()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        using (var zip = ZipFile.Open(source, ZipArchiveMode.Update))
        {
            var sheet = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            byte[] original;
            using (var input = sheet.Open())
            using (var memory = new MemoryStream())
            {
                input.CopyTo(memory);
                original = memory.ToArray();
            }
            sheet.Delete();
            var changed = Encoding.UTF8.GetString(original).Replace("<c r=\"B1\" t=\"n\">", "<c r=\"B1\" s=\"0\" t=\"n\">", StringComparison.Ordinal);
            Assert.NotEqual(Encoding.UTF8.GetString(original), changed);
            using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open();
            output.Write(Encoding.UTF8.GetBytes(changed));
        }
        using (var store = new PackageStore(source))
        {
            var result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "E5", "blank", null, Operation: "clear")]);
            Assert.Empty(result.ChangedParts);
            Assert.Empty(result.Edits);
            var untouched = fixture.Output("untouched");
            store.Save(untouched);
            Assert.Empty(P2aGates.CheckIntent(untouched, [new CellExpectation("Sheet1", "E5", "blank", null, true)]));
            Assert.Empty(P2aGates.ReadCells(untouched, "Sheet1", ["E5"]));
        }
        using var sessions = new ExcelSessions();
        var id = Open(sessions, source);
        sessions.Apply(id, 0, Request("""{"op":"clear","sheet":"Sheet1","target":"B1"}""").NormalizeMany(null));
        var path = fixture.Output("styled");
        sessions.Save(id, path);
        using var archive = ZipFile.OpenRead(path);
        using var inputSheet = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        Assert.Contains("r=\"B1\" s=\"0\"", inputSheet.ReadToEnd());
        Assert.Equal("blank", Assert.Single(P2aGates.ReadCells(path, "Sheet1", ["B1"])).Kind);
        sessions.Close(id, true);
    }

    [Fact]
    public void ClearingFormulaRemovesCalcChainButClearingOtherCellsDoesNot()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        using (var zip = ZipFile.Open(source, ZipArchiveMode.Update))
        {
            AddBeforeClose(zip, "xl/_rels/workbook.xml.rels", "</Relationships>",
                $"<Relationship Id=\"rId9\" Type=\"{office}/calcChain\" Target=\"calcChain.xml\"/>");
            AddBeforeClose(zip, "[Content_Types].xml", "</Types>",
                "<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>");
            using var chain = zip.CreateEntry("xl/calcChain.xml").Open();
            chain.Write(Encoding.UTF8.GetBytes("<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"C1\" i=\"1\"/></calcChain>"));
        }
        foreach (var (target, removed, removeCell) in new[] { ("B1", false, false), ("C1", true, false),
                     ("E5", false, false), ("C1", true, true) })
        {
            using var sessions = new ExcelSessions();
            var id = Open(sessions, source);
            sessions.Apply(id, 0, Request($"{{\"op\":\"clear\",\"sheet\":\"Sheet1\",\"target\":\"{target}\",\"remove_cells\":{removeCell.ToString().ToLowerInvariant()}}}").NormalizeMany(null));
            var path = fixture.Output(target + removeCell);
            sessions.Save(id, path);
            using var archive = ZipFile.OpenRead(path);
            Assert.Equal(removed, archive.GetEntry("xl/calcChain.xml") is null);
            if (removeCell) Assert.Empty(P2aGates.ReadCells(path, "Sheet1", [target]));
            sessions.Close(id, true);
        }
    }

    [Theory]
    [InlineData("{\"op\":\"clear\",\"target\":\"B1\",\"what\":[\"formats\"]}")]
    [InlineData("{\"op\":\"clear\",\"target\":\"B1\",\"what\":[\"values\",\"formats\"]}")]
    [InlineData("{\"op\":\"clear\",\"target\":\"B1\",\"what\":[]}")]
    [InlineData("{\"op\":\"clear\",\"target\":\"B1\",\"remove_cells\":1}")]
    [InlineData("{\"op\":\"clear\",\"target\":\"B1\",\"remove_cells\":null}")]
    [InlineData("{\"op\":\"clear\",\"target\":\"B1\",\"value\":null}")]
    [InlineData("{\"op\":\"clear\",\"target\":\"B1\",\"rich_policy\":\"replace\"}")]
    [InlineData("{\"op\":\"clear\",\"target\":\"B1\",\"unknown\":1}")]
    public void UnsupportedClearFacetsFailClosed(string json)
    {
        Assert.Throws<NotSupportedException>(() => Request(json).NormalizeMany("Sheet1"));
    }

    [Fact]
    public void EngineRejectsInconsistentClearOperation()
    {
        using var fixture = new Fixture();
        using var store = new PackageStore(fixture.Source("default"));
        Assert.Throws<NotSupportedException>(() => SetValueEngine.Apply(store,
            [new SetValueOp("Sheet1", "B1", "number", "7", Operation: "clear")]));
        Assert.Empty(store.ChangedParts);
        Assert.Throws<NotSupportedException>(() => SetValueEngine.Apply(store,
            [new SetValueOp("Sheet1", "B1", "blank", null, RemoveCell: true)]));
    }

    [Fact]
    public void LostClearAndRemovedExistingCellFailG4AndG5()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        Assert.Contains(P2aGates.CheckIntent(source, [new CellExpectation("Sheet1", "B1", "blank", null, true)]),
            issue => issue.Code == "INTENT_MISMATCH");
        var required = new CellExpectation("Sheet1", "B1", "blank", null, true, true);
        Assert.Contains(P2aGates.CheckIntent(source, [required]), issue => issue.Code == "INTENT_PRESENT");
        Assert.Contains(P2aGates.CheckTouchedCells(source, source, [required]),
            issue => issue.Code == "CELL_NOT_REMOVED");
        var missing = fixture.Output("missing");
        File.Copy(source, missing);
        using (var zip = ZipFile.Open(missing, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            byte[] bytes;
            using (var input = entry.Open())
            using (var memory = new MemoryStream())
            {
                input.CopyTo(memory);
                bytes = memory.ToArray();
            }
            entry.Delete();
            var text = System.Text.RegularExpressions.Regex.Replace(Encoding.UTF8.GetString(bytes),
                @"<c r=""B1"" t=""n"">\s*<v>42</v>\s*</c>", "");
            Assert.NotEqual(Encoding.UTF8.GetString(bytes), text);
            using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open();
            output.Write(Encoding.UTF8.GetBytes(text));
        }
        Assert.Contains(P2aGates.CheckTouchedCells(source, missing,
            [new CellExpectation("Sheet1", "B1", "blank", null, true)]), issue => issue.Code == "CELL_MISSING");
        Assert.Contains(P2aGates.CheckIntent(missing,
            [new CellExpectation("Sheet1", "B1", "blank", null)]), issue => issue.Code == "INTENT_MISSING");

        var created = fixture.Output("created");
        File.Copy(source, created);
        using (var zip = ZipFile.Open(created, ZipArchiveMode.Update))
            AddBeforeClose(zip, "xl/worksheets/sheet1.xml", "</sheetData>", "<row r=\"5\"><c r=\"E5\"/></row>");
        Assert.Contains(P2aGates.CheckTouchedCells(source, created,
            [new CellExpectation("Sheet1", "E5", "blank", null, true)]),
            issue => issue.Code == "UNEXPECTED_CELL_CREATED");
    }

    [Fact]
    public void G4RejectsStaleCellTypeAfterClear()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        var output = fixture.Output("stale-type");
        using (var store = new PackageStore(source))
        {
            SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "blank", null, Operation: "clear")]);
            store.Save(output);
        }
        using (var zip = ZipFile.Open(output, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            byte[] bytes;
            using (var input = entry.Open())
            using (var memory = new MemoryStream())
            {
                input.CopyTo(memory);
                bytes = memory.ToArray();
            }
            entry.Delete();
            var text = Encoding.UTF8.GetString(bytes).Replace("<c r=\"B1\"", "<c r=\"B1\" t=\"s\"", StringComparison.Ordinal);
            Assert.NotEqual(Encoding.UTF8.GetString(bytes), text);
            using var modified = zip.CreateEntry("xl/worksheets/sheet1.xml").Open();
            modified.Write(Encoding.UTF8.GetBytes(text));
        }
        Assert.Contains(P2aGates.CheckIntent(output,
            [new CellExpectation("Sheet1", "B1", "blank", null, true)]),
            issue => issue.Code == "INTENT_MISMATCH");
    }

    private static string Open(ExcelSessions sessions, string source) =>
        JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;

    private static SetValueRequest Request(string json) => JsonSerializer.Deserialize<SetValueRequest>(json)!;

    private static void AddBeforeClose(ZipArchive zip, string part, string closing, string addition)
    {
        var entry = zip.GetEntry(part)!;
        byte[] bytes;
        using (var input = entry.Open())
        using (var memory = new MemoryStream())
        {
            input.CopyTo(memory);
            bytes = memory.ToArray();
        }
        entry.Delete();
        using var output = zip.CreateEntry(part).Open();
        output.Write(Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace(closing, addition + closing, StringComparison.Ordinal)));
    }

    private sealed class Fixture : IDisposable
    {
        public string Directory { get; } = Path.Combine(Path.GetTempPath(), "docloupe-clear-" + Guid.NewGuid().ToString("N"));
        public string Source(string variant)
        {
            SyntheticFixtures.Create(Directory);
            return Path.Combine(Directory, variant + ".xlsx");
        }
        public string Output(string name) => Path.Combine(Directory, name + "-cleared.xlsx");
        public void Dispose() => System.IO.Directory.Delete(Directory, true);
    }
}
