using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SavePipelineTests
{
    [Fact]
    public void QualifiedAddressOverridesDefaultSheet()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-qualified-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var opened = sessions.Open(Path.Combine(directory, "default.xlsx"));
            var id = (string)opened.GetType().GetProperty("session")!.GetValue(opened)!;
            var operation = new SetValueRequest
            {
                Op = "set_value", Sheet = "WrongDefault", Target = "Sheet1!B1",
                Value = System.Text.Json.JsonSerializer.SerializeToElement(27)
            };
            sessions.Apply(id, 0, [operation.Normalize("OtherDefault")]);
            var preview = sessions.Read(id, "WrongDefault", ["Sheet1!B1"]);
            var previewCells = (IReadOnlyList<DocLoupe.Excel.Verify.CellRead>)preview.GetType().GetProperty("cells")!.GetValue(preview)!;
            Assert.Equal("27", Assert.Single(previewCells).Value);
            var output = Path.Combine(directory, "qualified.xlsx");
            sessions.Save(id, output);
            Assert.Equal("27", Assert.Single(DocLoupe.Excel.Verify.P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ExplicitFormulaLikeTextUsesSharedStringStorage()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-as-text-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var opened = sessions.Open(Path.Combine(directory, "default.xlsx"));
            var id = (string)opened.GetType().GetProperty("session")!.GetValue(opened)!;
            var operation = new SetValueRequest
            {
                Op = "set_value", Target = "B1", AsText = true,
                Value = System.Text.Json.JsonSerializer.SerializeToElement("=1+2")
            };
            sessions.Apply(id, 0, [operation.Normalize("Sheet1")]);
            var output = Path.Combine(directory, "formula-text.xlsx");
            sessions.Save(id, output);
            var cell = Assert.Single(DocLoupe.Excel.Verify.P2aGates.ReadCells(output, "Sheet1", ["B1"]));
            Assert.Equal("text", cell.Kind);
            Assert.Equal("=1+2", cell.Value);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void UndoReplaysOnlyRetainedRevisions()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-undo-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var opened = sessions.Open(Path.Combine(directory, "default.xlsx"));
            var id = (string)opened.GetType().GetProperty("session")!.GetValue(opened)!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            sessions.Apply(id, 1, [new SetValueOp("Sheet1", "B1", "number", "28"),
                new SetValueOp("Sheet1", "D3", "number", "8")]);
            Assert.Throws<InvalidOperationException>(() => sessions.Undo(id, 1, 0));
            var undone = sessions.Undo(id, 2, 1);
            Assert.Equal(1, undone.Revision);
            Assert.Equal([2], undone.Discarded);
            Assert.Throws<ArgumentOutOfRangeException>(() => sessions.Undo(id, 1, 2));
            sessions.Apply(id, 1, [new SetValueOp("Sheet1", "B1", "number", "29")]);
            var output = Path.Combine(directory, "undone.xlsx");
            sessions.Save(id, output);
            Assert.Equal("29", Assert.Single(DocLoupe.Excel.Verify.P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            Assert.Equal("old", Assert.Single(DocLoupe.Excel.Verify.P2aGates.ReadCells(output, "Sheet1", ["D3"])).Value);
            sessions.Undo(id, 2, 0);
            var preview = sessions.Read(id, "Sheet1", ["B1"]);
            var cells = (IReadOnlyList<DocLoupe.Excel.Verify.CellRead>)preview.GetType().GetProperty("cells")!.GetValue(preview)!;
            Assert.Equal("42", Assert.Single(cells).Value);
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

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
