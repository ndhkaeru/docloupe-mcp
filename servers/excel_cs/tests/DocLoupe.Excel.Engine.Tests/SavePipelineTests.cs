using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SavePipelineTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("nested-workbook")]
    public void SavesOnlyAfterAllFiveGates(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-save-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var opened = sessions.Open(Path.Combine(directory, variant + ".xlsx"));
            var id = (string)opened.GetType().GetProperty("session")!.GetValue(opened)!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            sessions.Apply(id, 1, [new SetValueOp("Sheet1", "B1", "number", "28")]);
            var output = Path.Combine(directory, "saved.xlsx");
            var report = sessions.Save(id, output);
            Assert.Equal("verified", report.GetType().GetProperty("status")!.GetValue(report));
            Assert.Equal("28", Assert.Single(DocLoupe.Excel.Verify.P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            Assert.True(File.Exists(output));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }
}
