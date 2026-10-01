using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SaveAssertionTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void G7MatchesValuesAndFormulasIndependentlyOfEdits(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-pass-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, variant + ".xlsx");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27"),
                new SetValueOp("Sheet1", "E5", "boolean", "true")]);
            var assertions = new[]
            {
                Assertion("Sheet1!B1", new { value = 27.0 }),
                Assertion("Sheet1!A1", new { value = "hello" }),
                Assertion("Sheet1!D3", new { value = "old" }),
                Assertion("Sheet1!E5", new { value = true }),
                Assertion("Sheet1!C1", new { formula = "=1+1" }),
                Assertion("Sheet1!F9", new { value = (object?)null })
            }.Select(item => item.Normalize()).ToArray();
            var output = Path.Combine(directory, "passed.xlsx");
            var result = JsonSerializer.SerializeToElement(sessions.Save(id, output, assertions));
            Assert.Contains(result.GetProperty("gates").EnumerateArray(), gate => gate.GetString() == "G7");
            Assert.Equal(assertions.Length, result.GetProperty("assertions").GetArrayLength());
            Assert.All(result.GetProperty("assertions").EnumerateArray(), assertion =>
                Assert.Equal("verified", assertion.GetProperty("status").GetString()));
            Assert.True(File.Exists(output));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G7BlocksWrongValueOrFormulaAndLeavesNoDestination()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-block-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            var output = Path.Combine(directory, "blocked.xlsx");
            var assertions = new[]
            {
                Assertion("Sheet1!B1", new { value = 28 }),
                Assertion("Sheet1!C1", new { formula = "9+9" }),
                Assertion("Wrong!A1", new { value = 1 })
            }.Select(item => item.Normalize()).ToArray();
            var blocked = Assert.Throws<SaveBlockedException>(() => sessions.Save(id, output, assertions));
            Assert.Contains(blocked.Issues, issue => issue.Gate == "G7" && issue.Code == "ASSERT_VALUE_MISMATCH");
            Assert.Contains(blocked.Issues, issue => issue.Gate == "G7" && issue.Code == "ASSERT_FORMULA_MISMATCH");
            Assert.Contains(blocked.Issues, issue => issue.Gate == "G7" && issue.Code == "ASSERT_SHEET_MISSING");
            Assert.False(File.Exists(output));
            Assert.Empty(Directory.GetFiles(directory, "*.staging"));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{\"display\":\"27\"}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{\"value\":27},\"unchanged\":true}")]
    [InlineData("{\"target\":\"B1\",\"equals\":{\"value\":27}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{\"value\":{\"date\":\"2026-10-15\"}}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{\"value\":1e2147483648}}")]
    public void RejectsUnsupportedOrAmbiguousAssertions(string json)
    {
        var assertion = JsonSerializer.Deserialize<SaveAssertionRequest>(json)!;
        var error = Record.Exception(() => assertion.Normalize());
        Assert.True(error is ArgumentException or NotSupportedException, error?.ToString());
    }

    [Fact]
    public void G7IndependentReaderDetectsCellMismatch()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-reader-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var path = Path.Combine(directory, "default.xlsx");
            Assert.Empty(G7Assertions.Check(path, [Assertion("Sheet1!B1", new { value = 42 }).Normalize()]));
            Assert.Contains(G7Assertions.Check(path, [Assertion("Sheet1!B1", new { value = 43 }).Normalize()]),
                issue => issue.Code == "ASSERT_VALUE_MISMATCH");
            Assert.Empty(G7Assertions.Check(path, [Assertion("Sheet1!B1", new { value = 4.2e1 }).Normalize()]));
            using (var store = new PackageStore(path))
            {
                SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "0")]);
                var zero = Path.Combine(directory, "zero.xlsx");
                store.Save(zero);
                Assert.Contains(G7Assertions.Check(zero, [Assertion("Sheet1!B1", new { value = 1e-100 }).Normalize()]),
                    issue => issue.Code == "ASSERT_VALUE_MISMATCH");
            }
            Assert.Contains(G7Assertions.Check(path, [Assertion("Sheet1!C1", new { value = 2 }).Normalize()]),
                issue => issue.Code == "ASSERT_CACHE_UNSUPPORTED");
        }
        finally { Directory.Delete(directory, true); }
    }

    private static SaveAssertionRequest Assertion(string target, object equals) => new()
    {
        Target = target, Expected = JsonSerializer.SerializeToElement(equals)
    };
}
