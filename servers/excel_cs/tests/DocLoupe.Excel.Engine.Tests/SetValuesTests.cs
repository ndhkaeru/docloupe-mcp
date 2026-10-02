using System.Text.Json;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SetValuesTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void RectangularBatchPassesSaveGates(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-values-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var operations = Request("""
                {"op":"set_values","target":"Sheet1!B1:C2","values":[[7,"hello"],[true,null]]}
                """).NormalizeMany("WrongDefault");
            Assert.Equal(4, operations.Length);
            Assert.All(operations, operation => Assert.Equal("set_values", operation.Operation));
            sessions.Apply(session, 0, operations);
            var status = JsonSerializer.SerializeToElement(sessions.Status(session));
            Assert.Equal(4, status.GetProperty("ledger")[0].GetProperty("op_count").GetInt32());
            var output = Path.Combine(directory, "result.xlsx");
            var saved = JsonSerializer.SerializeToElement(sessions.Save(session, output,
                [new ValueAssertion("Sheet1", "C1", true, "text", "hello", null)]));
            Assert.Equal("verified", saved.GetProperty("status").GetString());
            var cells = P2aGates.ReadCells(output, "Sheet1", ["B1", "C1", "B2", "C2"]);
            var byAddress = cells.ToDictionary(cell => cell.Address);
            Assert.Equal("7", byAddress["B1"].Value);
            Assert.Equal("hello", byAddress["C1"].Value);
            Assert.Equal("true", byAddress["B2"].Value);
            Assert.Null(byAddress["C2"].Value);
            sessions.Undo(session, 1, 0);
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void AnchorExpandsAndMixesWithSingleCellOperation()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-values-anchor-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var operations = new[]
            {
                Request("""{"op":"set_values","target":"D4","values":[[3,"item"]]}""").NormalizeMany("Sheet1"),
                Request("""{"op":"set_formula","target":"C1","formula":"D4+1"}""").NormalizeMany("Sheet1")
            }.SelectMany(batch => batch).ToArray();
            sessions.Apply(session, 0, operations);
            var output = Path.Combine(directory, "mixed.xlsx");
            sessions.Save(session, output);
            var cells = P2aGates.ReadCells(output, "Sheet1", ["D4", "E4", "C1"]);
            var byAddress = cells.ToDictionary(cell => cell.Address);
            Assert.Equal("3", byAddress["D4"].Value);
            Assert.Equal("item", byAddress["E4"].Value);
            Assert.Equal("D4+1", byAddress["C1"].Formula);
            sessions.Close(session, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("{\"op\":\"set_values\",\"target\":\"B1:C2\",\"values\":[[1,2]]}")]
    [InlineData("{\"op\":\"set_values\",\"target\":\"B1\",\"values\":[[1,2],[3]]}")]
    [InlineData("{\"op\":\"set_values\",\"target\":\"B1\",\"values\":[]}")]
    [InlineData("{\"op\":\"set_values\",\"target\":\"B1\",\"values\":[[]]}")]
    [InlineData("{\"op\":\"set_values\",\"target\":\"XFD1048576\",\"values\":[[1,2]]}")]
    [InlineData("{\"op\":\"set_values\",\"target\":\"C2:B1\",\"values\":[[1]]}")]
    [InlineData("{\"op\":\"set_values\",\"target\":\"B1\",\"values\":[[\"=1+2\"]]}")]
    [InlineData("{\"op\":\"set_values\",\"target\":\"B1\",\"values\":[[1]],\"as_text\":true}")]
    [InlineData("{\"op\":\"set_value\",\"target\":\"B1\",\"value\":1,\"values\":[[2]]}")]
    public void MalformedBatchCannotBeApplied(string json)
    {
        var error = Record.Exception(() => Request(json).NormalizeMany("Sheet1"));
        Assert.True(error is ArgumentException or FormatException or NotSupportedException, error?.ToString());
    }

    [Fact]
    public void InvalidBulkOpLeavesSessionUnchanged()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-values-atomic-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var requests = new[]
            {
                Request("""{"op":"set_value","target":"B1","value":99}"""),
                Request("""{"op":"set_values","target":"C1:D2","values":[[1,2]]}""")
            };
            Assert.Throws<ArgumentException>(() => sessions.Apply(session, 0,
                requests.SelectMany(request => request.NormalizeMany("Sheet1")).ToArray()));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(session)).GetProperty("revision").GetInt32());
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static SetValueRequest Request(string json) => JsonSerializer.Deserialize<SetValueRequest>(json)!;
}
