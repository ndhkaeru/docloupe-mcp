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

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void ExpectEmptyChecksContentAndFormulaWithoutUsingItsCache(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-expect-empty-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("""{"op":"set_value","target":"Sheet1!E5","value":7,"expect":{"empty":true,"value":null}}""")]);
            var failure = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 1,
            [
                Request("""{"op":"clear","target":"Sheet1!B1"}"""),
                Request("""{"op":"set_value","target":"Sheet1!E5","value":9,"expect":{"empty":true}}""")
            ]));
            Assert.Equal(1, failure.Index);
            Assert.True(failure.Expected.Empty);
            Assert.Equal("7", failure.Actual?.Value);
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            Assert.Equal("42", Assert.Single(P2aGates.ReadCells(Path.Combine(directory, variant + ".xlsx"), "Sheet1", ["B1"])).Value);

            sessions.Apply(id, 1, [Request("""{"op":"set_formula","target":"Sheet1!C1","formula":"=3+3","cache":"clear","expect":{"empty":false}}""")]);
            var formula = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 2,
                [Request("""{"op":"clear","target":"Sheet1!C1","expect":{"empty":true,"value":null}}""")]));
            Assert.Equal("3+3", formula.Actual?.Formula);
            sessions.Apply(id, 2, [Request("""{"op":"set_value","target":"Sheet1!B1","value":12,"expect":{"empty":false,"value":42}}""")]);
            sessions.Apply(id, 3, [Request("""{"op":"clear","target":"Sheet1!E5","expect":{"empty":false}}""")]);
            sessions.Apply(id, 4, [Request("""{"op":"set_value","target":"Sheet1!E5","value":8,"expect":{"empty":true}}""")], dryRun: true);
            Assert.Equal(4, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
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
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void BulkPreconditionsCheckEveryCellBeforeTheBatch(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-expect-bulk-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var initial = new[]
            {
                """{"op":"set_value","target":"Sheet1!E5:F5","value":7,"expect":{"empty":true}}""",
                """{"op":"fill","target":"Sheet1!E6:F6","series":{"start":1,"step":1},"expect":{"empty":true}}""",
                """{"op":"set_values","target":"Sheet1!E7","values":[[3,4]],"expect":{"empty":true}}"""
            }.SelectMany((json, index) => JsonSerializer.Deserialize<SetValueRequest>(json)!.NormalizeMany(null)
                .Select(cell => cell with { SourceIndex = index })).ToArray();
            Assert.Equal(6, initial.Length);
            Assert.All(initial, operation => Assert.True(operation.Expect?.Empty));
            sessions.Apply(id, 0, initial);

            var invalid = new[]
            {
                """{"op":"set_value","target":"Sheet1!I9:J9","value":9,"expect":{"empty":true}}""",
                """{"op":"fill","target":"Sheet1!B1:C1","series":{"start":5,"step":1},"expect":{"value":42}}"""
            }.SelectMany((json, index) => JsonSerializer.Deserialize<SetValueRequest>(json)!.NormalizeMany(null)
                .Select(cell => cell with { SourceIndex = index })).ToArray();
            var failure = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 1, invalid));
            Assert.Equal(1, failure.Index);
            Assert.Equal("Sheet1!C1", failure.Target);
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["I9", "J9"]))
                .GetProperty("cells").EnumerateArray());

            var clear = JsonSerializer.Deserialize<SetValueRequest>("""{"op":"clear","target":"Sheet1!B1:C1","expect":{"empty":false}}""")!
                .NormalizeMany(null);
            sessions.Apply(id, 1, clear);
            var reset = JsonSerializer.Deserialize<SetValueRequest>("""{"op":"set_values","target":"Sheet1!B1:C1","values":[[11,12]],"expect":{"empty":true}}""")!
                .NormalizeMany(null);
            sessions.Apply(id, 2, reset);
            var output = Path.Combine(directory, "verified.xlsx");
            sessions.Save(id, output);
            Assert.Equal(["11", "12"], P2aGates.ReadCells(output, "Sheet1", ["B1", "C1"])
                .Select(cell => cell.Value!).ToArray());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("new-shared-strings")]
    public void ExpectRichChecksRunFormattingBeforeBatchAndCurrentRevision(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-expect-rich-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var richInput = Path.Combine(directory, "rich-input.xlsx");
            RichFixture.Create(Path.Combine(directory, variant + ".xlsx"), richInput,
                variant == "new-shared-strings");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(richInput)).GetProperty("session").GetString()!;
            const string good = """{"op":"set_value","target":"Sheet1!A1","value":"replaced","rich_policy":"replace","expect":{"value":"hello bold","text":"hello bold","rich":"<r>hello</r><r b color=\"FF0000\"> bold</r>"}}""";
            const string wrong = """{"op":"set_value","target":"Sheet1!A1","value":"replaced","rich_policy":"replace","expect":{"rich":"<r>hello</r><r b=\"0\" color=\"FF0000\"> bold</r>"}}""";
            sessions.Apply(id, 0, [Request(good)], dryRun: true);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            var failed = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 0,
                [Request("""{"op":"set_value","target":"Sheet1!E5","value":99}"""), Request(wrong)]));
            Assert.Equal(1, failed.Index);
            Assert.Equal("Sheet1!A1", failed.Target);
            Assert.Contains("<r>hello</r>", failed.Expected.Rich);
            Assert.Equal("hello bold", failed.Actual?.Value);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["E5"]))
                .GetProperty("cells").EnumerateArray());
            var expanded = JsonSerializer.Deserialize<SetValueRequest>("""{"op":"set_value","target":"Sheet1!A1:B1","value":"new","rich_policy":"replace","expect":{"rich":"<r>hello</r><r b color=\"FF0000\"> bold</r>"}}""")!
                .NormalizeMany(null).Select(cell => cell with { SourceIndex = 1 }).ToArray();
            var rangeFailure = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 0,
                [Request("""{"op":"set_value","target":"Sheet1!E5","value":99}"""), .. expanded]));
            Assert.Equal(1, rangeFailure.Index);
            Assert.Equal("Sheet1!B1", rangeFailure.Target);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Apply(id, 0, [Request(good)]);
            var repeated = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 1, [Request(good)]));
            Assert.Equal("replaced", repeated.Actual?.Value);
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            var output = Path.Combine(directory, "verified.xlsx");
            sessions.Save(id, output);
            Assert.Equal("replaced", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["A1"])).Value);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("new-shared-strings")]
    public void ExpectExplicitBoldStyleChecksBeforeBatchAndDryRun(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-expect-style-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var styled = Path.Combine(directory, "styled.xlsx");
            StyledFixture.Create(Path.Combine(directory, variant + ".xlsx"), styled);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(styled)).GetProperty("session").GetString()!;
            const string good = """{"op":"set_value","target":"Sheet1!B1","value":27,"expect":{"value":42,"style":{"font":{"bold":true}}}}""";
            const string wrong = """{"op":"set_value","target":"Sheet1!B1","value":28,"expect":{"style":{"font":{"bold":false}}}}""";
            sessions.Apply(id, 0, [Request(good)], dryRun: true);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            var failed = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 0,
                [Request("""{"op":"set_value","target":"Sheet1!E5","value":99}"""), Request(wrong)]));
            Assert.Equal(1, failed.Index);
            Assert.Equal("Sheet1!B1", failed.Target);
            Assert.False(failed.Expected.FontBold);
            Assert.True(failed.ActualFontBold);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["E5"]))
                .GetProperty("cells").EnumerateArray());
            var range = JsonSerializer.Deserialize<SetValueRequest>("""{"op":"set_value","target":"Sheet1!B1:C1","value":30,"expect":{"style":{"font":{"bold":true}}}}""")!
                .NormalizeMany(null).Select(cell => cell with { SourceIndex = 1 }).ToArray();
            var rangeFailure = Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 0,
                [Request("""{"op":"set_value","target":"Sheet1!E5","value":99}"""), .. range]));
            Assert.Equal(1, rangeFailure.Index);
            Assert.Equal("Sheet1!C1", rangeFailure.Target);
            Assert.Null(rangeFailure.ActualFontBold);
            sessions.Apply(id, 0, [Request(good), Request("""{"op":"clear","target":"Sheet1!D3","expect":{"style":{"font":{"bold":false}},"empty":false}}""")]);
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Apply(id, 1, [Request("""{"op":"set_value","target":"Sheet1!B1","value":28,"expect":{"value":27,"style":{"font":{"bold":true}}}}""")]);
            var output = Path.Combine(directory, "verified.xlsx");
            sessions.Save(id, output);
            Assert.Equal("28", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":null}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"formula":""}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"empty":null}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"empty":1}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"empty":true,"empty":false}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"empty":true,"display":"42"}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"formula":7}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"formula":"=1+1","value":{}}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"formula":"1+1","formula":"2+2"}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"value":42,"rich":"x"}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!A1","value":"x","expect":{"rich":"<r bad=\"1\">x</r>"}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!A1","value":"x","expect":{"rich":null}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!A1","value":"x","expect":{"rich":"x","empty":true}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"style":{}}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"style":{"font":{"bold":1}}}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"style":{"font":{"bold":true,"italic":true}}}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1","value":1,"expect":{"value":1e2147483648}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!B1:C1","value":1,"expect":{"rich":17}}""")]
    [InlineData("""{"op":"set_values","target":"Sheet1!B1","values":[[1]],"expect":{"style":{}}}""")]
    [InlineData("""{"op":"fill","target":"Sheet1!B1:C1","value":1,"expect":{"empty":null}}""")]
    public void UnsupportedPreconditionShapesFailClosed(string json)
    {
        var error = Record.Exception(() => Request(json));
        Assert.True(error is ArgumentException or NotSupportedException, error?.ToString());
    }

    private static DocLoupe.Excel.Engine.SetValueOp Request(string json) =>
        Assert.Single(JsonSerializer.Deserialize<SetValueRequest>(json)!.NormalizeMany(null));
}
