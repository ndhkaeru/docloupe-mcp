using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ReadValuesTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void ReadsTypedRectanglesAndBlanksFromCurrentRevision(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-values-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var original = JsonSerializer.SerializeToElement(sessions.Read(id, null, ["Sheet1!A1:C2"], view: "values"));
            var rows = original.GetProperty("rows");
            Assert.Equal("values", original.GetProperty("view").GetString());
            Assert.Equal("hello", rows[0][0].GetString());
            Assert.Equal(42, rows[0][1].GetInt32());
            Assert.Equal(2, rows[0][2].GetInt32());
            Assert.All(rows[1].EnumerateArray(), cell => Assert.Equal(JsonValueKind.Null, cell.ValueKind));

            sessions.Apply(id, 0,
            [
                new SetValueOp("Sheet1", "A2", "number", "0.125"),
                new SetValueOp("Sheet1", "B2", "boolean", "true"),
                new SetValueOp("Sheet1", "C2", "error", "#N/A"),
                new SetValueOp("Sheet1", "D2", "formula", "\"done\"", Operation: "set_formula",
                    ExplicitCache: new FormulaCache("str", "done"))
            ]);
            var updated = JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["A2:D2"], view: "values"));
            var values = updated.GetProperty("rows")[0];
            Assert.Equal(1, updated.GetProperty("revision").GetInt32());
            Assert.Equal("0.125", values[0].GetRawText());
            Assert.True(values[1].GetBoolean());
            Assert.Equal("#N/A", values[2].GetProperty("error").GetString());
            Assert.Equal("done", values[3].GetString());
            Assert.Equal(JsonValueKind.Null,
                JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["Z99"], view: "values"))
                    .GetProperty("rows")[0][0].ValueKind);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ValuesViewRejectsNonRectangularRequestsAndUnknownViews()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-values-shape-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.Throws<NotSupportedException>(() => sessions.Read(id, "Sheet1", ["A1", "C1"], view: "values"));
            Assert.Throws<NotSupportedException>(() => sessions.Read(id, "Sheet1", ["A1"], view: "full"));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }
}
