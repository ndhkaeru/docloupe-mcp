using System.Text.Json;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SetValueRangeTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void BroadcastsValueToRangeAndPassesSaveGates(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-broadcast-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var operations = Request("""{"op":"set_value","target":"Sheet1!B1:C2","value":17}""").NormalizeMany("WrongDefault");
            Assert.Equal(["B1", "C1", "B2", "C2"], operations.Select(operation => operation.Address).ToArray());
            sessions.Apply(session, 0, operations);
            Assert.Equal(4, JsonSerializer.SerializeToElement(sessions.Status(session))
                .GetProperty("ledger")[0].GetProperty("op_count").GetInt32());
            var output = Path.Combine(directory, "result.xlsx");
            var saved = JsonSerializer.SerializeToElement(sessions.Save(session, output));
            Assert.Equal("verified", saved.GetProperty("status").GetString());
            var cells = P2aGates.ReadCells(output, "Sheet1", ["B1", "C1", "B2", "C2"]);
            Assert.Equal(4, cells.Count);
            Assert.All(cells, cell => Assert.Equal("17", cell.Value));
            sessions.Undo(session, 1, 0);
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void BroadcastPreservesExplicitFormulaLikeText()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-broadcast-text-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var operations = Request("""{"op":"set_value","target":"E4:F4","value":"=1+2","as_text":true}""")
                .NormalizeMany("Sheet1");
            sessions.Apply(session, 0, operations);
            var output = Path.Combine(directory, "text.xlsx");
            sessions.Save(session, output);
            Assert.All(P2aGates.ReadCells(output, "Sheet1", ["E4", "F4"]),
                cell => Assert.Equal("=1+2", cell.Value));
            sessions.Close(session, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("{\"op\":\"set_value\",\"target\":\"B1:A1\",\"value\":1}")]
    [InlineData("{\"op\":\"set_value\",\"target\":\"A1:ZZ1000\",\"value\":1}")]
    [InlineData("{\"op\":\"set_value\",\"target\":\"XFD1048576:XFE1048576\",\"value\":1}")]
    [InlineData("{\"op\":\"set_value\",\"target\":\"B1:C1\",\"value\":\"=1+2\"}")]
    [InlineData("{\"op\":\"set_value\",\"target\":\"B1:C1\"}")]
    [InlineData("{\"op\":\"set_value\",\"target\":\"B1:C1\",\"value\":1,\"values\":[[1,2]]}")]
    public void InvalidBroadcastFailsBeforeMutation(string json)
    {
        var error = Record.Exception(() => Request(json).NormalizeMany("Sheet1"));
        Assert.True(error is ArgumentException or FormatException or NotSupportedException, error?.ToString());
    }

    [Fact]
    public void RichCellInBroadcastKeepsRevisionAtomic()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-broadcast-rich-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var operations = Request("""{"op":"set_value","target":"A1:B1","value":"new"}""")
                .NormalizeMany("Sheet1");
            Assert.Throws<InvalidDataException>(() => sessions.Apply(session, 0, operations));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(session)).GetProperty("revision").GetInt32());
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static SetValueRequest Request(string json) => JsonSerializer.Deserialize<SetValueRequest>(json)!;
}
