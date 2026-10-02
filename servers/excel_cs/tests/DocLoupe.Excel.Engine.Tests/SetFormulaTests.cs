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
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"cache\":\"keep\"}")]
    [InlineData("{\"op\":\"set_formula\",\"target\":\"C1\",\"formula\":\"1+1\",\"cache\":{\"value\":2}}")]
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
