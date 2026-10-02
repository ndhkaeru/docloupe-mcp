using System.Text.Json;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ApplyPreconditionTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void ExpectValueUsesCurrentRevisionAndFailureDoesNotApplyBatch(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-expect-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("""{"op":"set_value","target":"Sheet1!B1","value":27,"expect":{"value":42.0}}""")]);
            sessions.Apply(id, 1, [Request("""{"op":"set_value","target":"Sheet1!B1","value":28,"expect":{"value":27}}""")]);

            var failure = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 2,
            [
                Request("""{"op":"set_value","target":"Sheet1!E5","value":99}"""),
                Request("""{"op":"set_value","target":"Sheet1!B1","value":30,"expect":{"value":42}}""")
            ]));
            Assert.Equal(1, failure.Index);
            Assert.Equal("Sheet1!B1", failure.Target);
            Assert.Equal("number", failure.Expected.Kind);
            Assert.Equal("28", failure.Actual?.Value);
            Assert.Equal(2, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["E5"]))
                .GetProperty("cells").EnumerateArray());

            var output = Path.Combine(directory, "verified.xlsx");
            sessions.Save(id, output);
            Assert.Equal("28", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            Assert.Empty(P2aGates.ReadCells(output, "Sheet1", ["E5"]));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ExpectValueChecksMissingAndFormulaCacheWithoutEvaluatingFormula()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-expect-typed-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("""{"op":"set_value","target":"Sheet1!E5","value":true,"expect":{"value":null}}""")]);
            var missing = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 1,
                [Request("""{"op":"clear","target":"Sheet1!E5","expect":{"value":null}}""")]));
            Assert.Equal("boolean", missing.Actual?.Kind);
            var notFormula = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 1,
                [Request("""{"op":"clear","target":"Sheet1!E5","expect":{"formula":"1+1"}}""")]));
            Assert.Null(notFormula.Actual?.Formula);
            sessions.Apply(id, 1, [Request("""{"op":"set_formula","target":"Sheet1!C1","formula":"=2+2","expect":{"value":2},"cache":{"value":4}}""")]);
            var formula = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 2,
                [Request("""{"op":"set_value","target":"Sheet1!C1","value":5,"expect":{"value":2}}""")]));
            Assert.Equal("4", formula.Actual?.Value);
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
    public void ExpectFormulaAndCacheCompareAgainstCurrentRevision(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-expect-formula-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("""{"op":"set_formula","target":"Sheet1!C1","formula":"=2+2","cache":{"value":4},"expect":{"formula":"=1+1","value":2}}""")]);
            var stale = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 1,
            [
                Request("""{"op":"set_value","target":"Sheet1!E5","value":7}"""),
                Request("""{"op":"clear","target":"Sheet1!C1","expect":{"formula":"=1+1"}}""")
            ]));
            Assert.Equal(1, stale.Index);
            Assert.Equal("1+1", stale.Expected.Formula);
            Assert.Equal("2+2", stale.Actual?.Formula);
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["E5"]))
                .GetProperty("cells").EnumerateArray());

            var wrongCache = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 1,
                [Request("""{"op":"set_value","target":"Sheet1!C1","value":9,"expect":{"formula":"2+2","value":2}}""")]));
            Assert.Equal("2+2", wrongCache.Actual?.Formula);
            Assert.Equal("4", wrongCache.Actual?.Value);
            sessions.Apply(id, 1, [Request("""{"op":"set_value","target":"Sheet1!C1","value":9,"expect":{"formula":"=2+2","value":4}}""")]);
            var output = Path.Combine(directory, "verified.xlsx");
            sessions.Save(id, output);
            Assert.Equal("9", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["C1"])).Value);
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
    public void ExpandedRangeReportsOriginalRequestIndexWithoutApplyingAnyCells(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-expect-range-index-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var requests = new[]
            {
                JsonSerializer.Deserialize<SetValueRequest>("""{"op":"set_value","target":"Sheet1!F5:G5","value":7}""")!,
                JsonSerializer.Deserialize<SetValueRequest>("""{"op":"clear","target":"Sheet1!B1","expect":{"value":0}}""")!
            };
            var expanded = requests.SelectMany((request, index) => request.NormalizeMany(null)
                .Select(cell => cell with { SourceIndex = index })).ToArray();
            Assert.Equal(3, expanded.Length);
            var failure = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 0, expanded));
            Assert.Equal(1, failure.Index);
            Assert.Equal("Sheet1!B1", failure.Target);
            Assert.Equal("42", failure.Actual?.Value);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["F5", "G5"]))
                .GetProperty("cells").EnumerateArray());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":null}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"formula":""}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"formula":7}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"formula":"=1+1","value":{}}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"formula":"1+1","formula":"2+2"}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"value":42,"rich":"x"}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"value":1e2147483648}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1:C1","value":1,"expect":{"value":42}}""")]
    [InlineData("""{"op":"set_values","target":"Sheet1!B1","values":[[1]],"expect":{"value":42}}""")]
    [InlineData("""{"op":"fill","target":"Sheet1!B1:C1","value":1,"expect":{"value":42}}""")]
    public void UnsupportedPreconditionShapesFailClosed(string json)
    {
        var error = Record.Exception(() => Request(json));
        Assert.True(error is ArgumentException or NotSupportedException, error?.ToString());
    }

    private static DocLoupe.Excel.Engine.SetValueOp Request(string json) =>
        Assert.Single(JsonSerializer.Deserialize<SetValueRequest>(json)!.NormalizeMany(null));
}
