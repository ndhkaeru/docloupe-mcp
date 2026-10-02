using System.IO.Compression;
using System.Text;
using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SetFormulaTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void NormalFormulaClearsCacheAndPassesSaveGates(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var operation = Request("{\"op\":\"set_formula\",\"target\":\"Sheet1!C1\",\"formula\":\"=B1+2\",\"cache\":\"clear\"}");
            sessions.Apply(id, 0, [operation.Normalize("WrongDefault")]);
            var status = JsonSerializer.SerializeToElement(sessions.Status(id));
            Assert.Contains("set_formula Sheet1!C1", status.GetProperty("ledger")[0].GetProperty("summary").GetString());
            var output = Path.Combine(directory, "formula.xlsx");
            var report = JsonSerializer.SerializeToElement(sessions.Save(id, output,
                [new ValueAssertion("Sheet1", "C1", false, null, null, "B1+2")]));
            Assert.Equal("verified", report.GetProperty("status").GetString());
            var read = Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["C1"]));
            Assert.Equal("B1+2", read.Formula);
            Assert.Null(read.Value);
            using var store = new PackageStore(output);
            var sheet = PackageStore.Parse(store.Read(store.SheetPart("Sheet1")));
            var cell = Assert.Single(sheet.GetElementsByTagName("c", PackageStore.Main).OfType<System.Xml.XmlElement>(),
                element => element.GetAttribute("r") == "C1");
            Assert.Empty(cell.GetElementsByTagName("v", PackageStore.Main).OfType<System.Xml.XmlElement>());
            var workbook = PackageStore.Parse(store.Read(store.WorkbookPart));
            var calc = Assert.Single(workbook.GetElementsByTagName("calcPr", PackageStore.Main).OfType<System.Xml.XmlElement>());
            Assert.Equal("1", calc.GetAttribute("fullCalcOnLoad"));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void KeepCachePreservesExistingFormulaResultAndPassesSaveGates(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-keep-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, variant + ".xlsx");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            var operation = Request("""{"op":"set_formula","target":"C1","formula":"=B1+2","cache":"keep"}""").Normalize("Sheet1");
            Assert.True(operation.KeepCache);
            sessions.Apply(id, 0, [operation]);
            var output = Path.Combine(directory, "kept.xlsx");
            sessions.Save(id, output, [new ValueAssertion("Sheet1", "C1", false, null, null, "B1+2")]);
            var cell = Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["C1"]));
            Assert.Equal("B1+2", cell.Formula);
            Assert.Equal("2", cell.Value);
            Assert.Empty(P2aGates.CheckTouchedCells(source, output,
                [new CellExpectation("Sheet1", "C1", "formula", "B1+2", KeepCache: true)]));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("b", "1", "true")]
    [InlineData("e", "#N/A", "#N/A")]
    [InlineData("str", "old", "old")]
    public void KeepCacheRetainsExistingResultTypes(string type, string cached, string expectedValue)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-cache-types-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
                entry.Delete();
                var changed = xml.Replace("<c r=\"C1\">", $"<c r=\"C1\" t=\"{type}\">", StringComparison.Ordinal)
                    .Replace("<v>2</v>", $"<v>{cached}</v>", StringComparison.Ordinal);
                Assert.NotEqual(xml, changed);
                using var output = archive.CreateEntry("xl/worksheets/sheet1.xml").Open();
                output.Write(Encoding.UTF8.GetBytes(changed));
            }
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("""{"op":"set_formula","target":"C1","formula":"B1+2","cache":"keep"}""")
                .Normalize("Sheet1")]);
            var outputPath = Path.Combine(directory, "kept.xlsx");
            sessions.Save(id, outputPath);
            Assert.Equal(expectedValue, Assert.Single(P2aGates.ReadCells(outputPath, "Sheet1", ["C1"])).Value);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("2.5", "", "2.5", "B1")]
    [InlineData("true", "b", "true", "C1")]
    [InlineData("false", "b", "false", "E5")]
    [InlineData("\"done & ready\"", "str", "done & ready", "C1")]
    [InlineData("{\"error\":\"#N/A\"}", "e", "#N/A", "C1")]
    public void ExplicitCacheWritesTypedResultAndPassesSaveGates(string value, string type, string result, string address)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-explicit-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            var operation = Request($"{{\"op\":\"set_formula\",\"target\":\"{address}\",\"formula\":\"1+1\",\"cache\":{{\"value\":{value}}}}}").Normalize("Sheet1");
            Assert.NotNull(operation.ExplicitCache);
            sessions.Apply(id, 0, [operation]);
            var output = Path.Combine(directory, "explicit.xlsx");
            var report = JsonSerializer.SerializeToElement(sessions.Save(id, output));
            Assert.Equal("verified", report.GetProperty("status").GetString());
            var read = Assert.Single(P2aGates.ReadCells(output, "Sheet1", [address]));
            Assert.Equal("1+1", read.Formula);
            Assert.Equal(result, read.Value);
            using var store = new PackageStore(output);
            var sheet = PackageStore.Parse(store.Read(store.SheetPart("Sheet1")));
            var cell = Assert.Single(sheet.GetElementsByTagName("c", PackageStore.Main).OfType<System.Xml.XmlElement>(),
                element => element.GetAttribute("r") == address);
            Assert.Equal(type, cell.GetAttribute("t"));
            Assert.Single(cell.GetElementsByTagName("v", PackageStore.Main).OfType<System.Xml.XmlElement>());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("new-shared-strings")]
    public void ExplicitCacheHandlesWorksheetVariants(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-explicit-variant-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("""{"op":"set_formula","target":"C1","formula":"1+3","cache":{"value":4}}""").Normalize("Sheet1")]);
            var output = Path.Combine(directory, "explicit.xlsx");
            Assert.Equal("verified", JsonSerializer.SerializeToElement(sessions.Save(id, output)).GetProperty("status").GetString());
            Assert.Equal("4", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["C1"])).Value);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G4RejectsMissingOrIncorrectExplicitCache()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-explicit-corrupt-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            foreach (var expected in new[] { new FormulaCacheExpectation("n", "3"), new FormulaCacheExpectation("b", "2") })
                Assert.Contains(P2aGates.CheckIntent(source,
                    [new CellExpectation("Sheet1", "C1", "formula", "1+1", ExplicitCache: expected)]),
                    issue => issue.Code == "INTENT_CACHE_MISMATCH");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G5RejectsExplicitCacheCorruption()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-explicit-g5-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            var output = Path.Combine(directory, "explicit.xlsx");
            using (var store = new PackageStore(source))
            {
                SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "C1", "formula", "1+3",
                    Operation: "set_formula", ExplicitCache: new FormulaCache("n", "4"))]);
                store.Save(output);
            }
            var intent = new CellExpectation("Sheet1", "C1", "formula", "1+3",
                ExplicitCache: new FormulaCacheExpectation("n", "9"));
            Assert.Contains(P2aGates.CheckTouchedCells(source, output, [intent]),
                issue => issue.Code == "FORMULA_CACHE_MISMATCH");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void InvalidExplicitCacheDoesNotAdvanceRevision()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-explicit-invalid-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var invalid = Request("""{"op":"set_formula","target":"E5","formula":"1+1","cache":{"value":1e999}}""").Normalize("Sheet1");
            Assert.Throws<FormatException>(() => sessions.Apply(id, 0,
                [Request("""{"op":"set_value","target":"B1","value":5}""").Normalize("Sheet1"), invalid]));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G5RejectsExplicitCacheCorruptionOnNewCell()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-explicit-new-g5-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            var output = Path.Combine(directory, "explicit.xlsx");
            using (var store = new PackageStore(source))
            {
                SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "E5", "formula", "1+3",
                    Operation: "set_formula", ExplicitCache: new FormulaCache("n", "4"))]);
                store.Save(output);
            }
            Assert.Contains(P2aGates.CheckTouchedCells(source, output,
                [new CellExpectation("Sheet1", "E5", "formula", "1+3",
                    ExplicitCache: new FormulaCacheExpectation("n", "5"))]),
                issue => issue.Code == "FORMULA_CACHE_MISMATCH");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G4RejectsUnclearedFormulaCache()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-clear-corrupt-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            Assert.Contains(P2aGates.CheckIntent(source,
                [new CellExpectation("Sheet1", "C1", "formula", "1+1")]),
                issue => issue.Code == "INTENT_CACHE_PRESENT");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void KeepCacheRejectsMissingOrNonFormulaSource()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-keep-guard-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            foreach (var target in new[] { "B1", "E5" })
            {
                var operation = Request($"{{\"op\":\"set_formula\",\"target\":\"{target}\",\"formula\":\"1+2\",\"cache\":\"keep\"}}").Normalize("Sheet1");
                Assert.Throws<NotSupportedException>(() => sessions.Apply(id, 0, [operation]));
                Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            }
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G5RejectsCacheLostAfterFormulaEdit()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-keep-corrupt-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            var output = Path.Combine(directory, "corrupt.xlsx");
            using (var store = new PackageStore(source))
            {
                SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "C1", "formula", "B1+2",
                    Operation: "set_formula", KeepCache: true)]);
                store.Save(output);
            }
            using (var archive = ZipFile.Open(output, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
                entry.Delete();
                var changed = xml.Replace("<v>2</v>", "<v>9</v>", StringComparison.Ordinal);
                Assert.NotEqual(xml, changed);
                using var stream = archive.CreateEntry("xl/worksheets/sheet1.xml").Open();
                stream.Write(Encoding.UTF8.GetBytes(changed));
            }
            var intent = new CellExpectation("Sheet1", "C1", "formula", "B1+2", KeepCache: true);
            Assert.Empty(P2aGates.CheckIntent(output, [intent]));
            Assert.Contains(P2aGates.CheckTouchedCells(source, output, [intent]),
                issue => issue.Code == "FORMULA_CACHE_CHANGED");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void MixedBatchAndUndoKeepRevisionAtomic()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-batch-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("{\"op\":\"set_value\",\"target\":\"B1\",\"value\":11}").Normalize("Sheet1"),
                Request("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"B1+1\"}").Normalize("Sheet1")]);
            var preview = JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["C1"]));
            Assert.Equal("B1+1", preview.GetProperty("cells")[0].GetProperty("Formula").GetString());
            var ledger = JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("ledger");
            Assert.Equal(2, ledger[0].GetProperty("op_count").GetInt32());
            sessions.Undo(id, 1, 0);
            var original = JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["C1"]));
            Assert.Equal("1+1", original.GetProperty("cells")[0].GetProperty("Formula").GetString());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("shared")]
    [InlineData("array")]
    [InlineData("dataTable")]
    public void ExistingFormulaGroupsCannotBeSilentlyDetached(string groupKind)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-formula-group-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
                entry.Delete();
                xml = xml.Replace("<f>1+1</f>", $"<f t=\"{groupKind}\" si=\"0\">1+1</f>", StringComparison.Ordinal);
                using var stream = archive.CreateEntry("xl/worksheets/sheet1.xml").Open();
                stream.Write(Encoding.UTF8.GetBytes(xml));
            }
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            Assert.Throws<NotSupportedException>(() => sessions.Apply(id, 0,
                [Request("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+3\"}").Normalize("Sheet1")]));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"kind\":\"shared\"}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"kind\":\"array\"}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"cache\":\"fresh\"}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"cache\":{\"value\":null}}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"cache\":{\"value\":{\"error\":\"bad\"}}}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"cache\":{\"value\":2,\"extra\":1}}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"cache\":{}}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"ref\":\"C1:C3\"}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"=\"}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"value\":5}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"rich_policy\":\"replace\"}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"si\":3}")]
    [InlineData("{\"op\":\"set_value\",\"target\":\"C1\",\"value\":1,\"formula\":\"1+1\"}")]
    public void UnsupportedFormulaFormsFailClosed(string json)
    {
        var error = Record.Exception(() => Request(json).Normalize("Sheet1"));
        Assert.True(error is ArgumentException or NotSupportedException, error?.ToString());
    }

    private static SetValueRequest Request(string json) => JsonSerializer.Deserialize<SetValueRequest>(json)!;
}
