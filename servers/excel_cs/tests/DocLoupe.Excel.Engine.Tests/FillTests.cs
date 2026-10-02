using System.Text.Json;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class FillTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void ConstantFillPassesSaveGates(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-fill-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var operations = Request("""{"op":"fill","target":"Sheet1!B1:C2","value":9}""").NormalizeMany("WrongDefault");
            Assert.Equal(4, operations.Length);
            Assert.All(operations, operation => Assert.Equal("fill", operation.Operation));
            sessions.Apply(session, 0, operations);
            var status = JsonSerializer.SerializeToElement(sessions.Status(session));
            Assert.Contains("fill Sheet1!B1", status.GetProperty("ledger")[0].GetProperty("summary").GetString());
            Assert.Equal(4, status.GetProperty("ledger")[0].GetProperty("op_count").GetInt32());
            var output = Path.Combine(directory, "result.xlsx");
            var saved = JsonSerializer.SerializeToElement(sessions.Save(session, output,
                [new ValueAssertion("Sheet1", "C2", true, "number", "9", null)]));
            Assert.Equal("verified", saved.GetProperty("status").GetString());
            var cells = P2aGates.ReadCells(output, "Sheet1", ["B1", "C1", "B2", "C2"]);
            Assert.Equal(4, cells.Count);
            Assert.All(cells, cell => Assert.Equal("9", cell.Value));
            sessions.Undo(session, 1, 0);
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void FormulaFillClearsCachesAndPassesReadback()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-fill-formula-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(session, 0, Request("""{"op":"fill","target":"B1:C1","value":{"formula":"=A1+2"}}""")
                .NormalizeMany("Sheet1"));
            var output = Path.Combine(directory, "formula.xlsx");
            sessions.Save(session, output,
                [new ValueAssertion("Sheet1", "C1", false, null, null, "A1+2")]);
            var cells = P2aGates.ReadCells(output, "Sheet1", ["B1", "C1"]);
            Assert.All(cells, cell =>
            {
                Assert.Equal("A1+2", cell.Formula);
                Assert.Null(cell.Value);
            });
            sessions.Close(session, true);
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
    public void IntegerSeriesFillsInRowMajorOrderAndPassesSaveGates(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-fill-series-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var operations = Request("""{"op":"fill","target":"Sheet1!B1:C2","series":{"start":-3,"step":2}}""")
                .NormalizeMany("WrongDefault");
            Assert.Equal(["B1", "C1", "B2", "C2"], operations.Select(operation => operation.Address).ToArray());
            Assert.Equal(["-3", "-1", "1", "3"], operations.Select(operation => operation.Value!).ToArray());
            Assert.All(operations, operation => Assert.Equal("fill", operation.Operation));
            sessions.Apply(session, 0, operations);
            var output = Path.Combine(directory, "series.xlsx");
            var saved = JsonSerializer.SerializeToElement(sessions.Save(session, output,
                [new ValueAssertion("Sheet1", "C2", true, "number", "3", null)]));
            Assert.Equal("verified", saved.GetProperty("status").GetString());
            var cells = P2aGates.ReadCells(output, "Sheet1", ["B1", "C1", "B2", "C2"])
                .ToDictionary(cell => cell.Address);
            Assert.Equal(4, cells.Count);
            Assert.Equal("-3", cells["B1"].Value);
            Assert.Equal("-1", cells["C1"].Value);
            Assert.Equal("1", cells["B2"].Value);
            Assert.Equal("3", cells["C2"].Value);
            sessions.Undo(session, 1, 0);
            sessions.Close(session, false);
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
    public void DecimalSeriesIsExactAndPassesSaveGates(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-fill-decimal-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var operations = Request("""{"op":"fill","target":"Sheet1!B1:C2","series":{"start":0.1,"step":0.2}}""")
                .NormalizeMany(null);
            Assert.Equal(["0.1", "0.3", "0.5", "0.7"], operations.Select(operation => operation.Value!).ToArray());
            sessions.Apply(session, 0, operations);
            var output = Path.Combine(directory, "decimal.xlsx");
            sessions.Save(session, output, [new ValueAssertion("Sheet1", "C2", true, "number", "0.7", null)]);
            Assert.Equal(["0.1", "0.3", "0.5", "0.7"], P2aGates.ReadCells(output, "Sheet1", ["B1", "C1", "B2", "C2"])
                .Select(cell => cell.Value!).ToArray());
            sessions.Close(session, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ScientificAndNegativeDecimalSeriesStayExact()
    {
        var operations = Request("""{"op":"fill","target":"B1:C2","series":{"start":3e2,"step":-0.1}}""")
            .NormalizeMany("Sheet1");
        Assert.Equal(["300", "299.9", "299.8", "299.7"], operations.Select(operation => operation.Value!).ToArray());
        var tiny = Request("""{"op":"fill","target":"B1:C1","series":{"start":1e-30,"step":1e-30}}""")
            .NormalizeMany("Sheet1");
        Assert.Equal("0." + new string('0', 29) + "1", tiny[0].Value);
        Assert.Equal("0." + new string('0', 29) + "2", tiny[1].Value);
    }

    [Theory]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":1.000000000000001,\"step\":0}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":1e-31,\"step\":0}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":1e30,\"step\":0}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":1e-30,\"step\":1}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":999999999999999.9,\"step\":0}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":1e999,\"step\":1}}")]
    public void UnsafeDecimalSeriesIsRejectedBeforeMutation(string json)
    {
        var error = Record.Exception(() => Request(json).NormalizeMany("Sheet1"));
        Assert.True(error is ArgumentException or FormatException or NotSupportedException, error?.ToString());
    }

    [Fact]
    public void BoundedIntegerSeriesAllowsZeroStep()
    {
        var operations = Request("""{"op":"fill","target":"Sheet1!B1:C1","series":{"start":999999999999999,"step":0}}""")
            .NormalizeMany(null);
        Assert.Equal(["999999999999999", "999999999999999"], operations.Select(operation => operation.Value!).ToArray());
    }

    [Theory]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":1,\"step\":1,\"extra\":0}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":1,\"start\":2,\"step\":1}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":9223372036854775807,\"step\":1}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":999999999999999,\"step\":1}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"series\":{\"start\":1,\"step\":9223372036854775807}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"value\":1,\"series\":{\"start\":1,\"step\":1}}")]
    [InlineData("{\"op\":\"set_value\",\"target\":\"B1:C1\",\"value\":1,\"series\":{\"start\":1,\"step\":1}}")]
    [InlineData("{\"op\":\"set_values\",\"target\":\"B1:C1\",\"values\":[[1,2]],\"series\":{\"start\":1,\"step\":1}}")]
    public void InvalidSeriesIsRejectedBeforeMutation(string json)
    {
        var error = Record.Exception(() => Request(json).NormalizeMany("Sheet1"));
        Assert.True(error is ArgumentException or FormatException or NotSupportedException, error?.ToString());
    }

    [Theory]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1\",\"value\":1}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\"}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"pattern_from\":\"A1\"}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"value\":1,\"series\":{\"start\":1,\"step\":2}}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"value\":1,\"as_text\":true}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"B1:C1\",\"value\":1,\"values\":[[1,2]]}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"C1:B1\",\"value\":1}")]
    [InlineData("{\"op\":\"fill\",\"target\":\"A1:ZZ1000\",\"value\":1}")]
    public void UnsupportedFillShapesFailClosed(string json)
    {
        var error = Record.Exception(() => Request(json).NormalizeMany("Sheet1"));
        Assert.True(error is ArgumentException or FormatException or NotSupportedException, error?.ToString());
    }

    [Fact]
    public void ExistingRichCellBlocksEntireFillRevision()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-fill-rich-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var operations = Request("""{"op":"fill","target":"A1:B1","value":"new"}""")
                .NormalizeMany("Sheet1");
            Assert.Throws<InvalidDataException>(() => sessions.Apply(session, 0, operations));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(session)).GetProperty("revision").GetInt32());
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static SetValueRequest Request(string json) => JsonSerializer.Deserialize<SetValueRequest>(json)!;
}
