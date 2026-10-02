using System.IO.Compression;
using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Model;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ErrorValueTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void ErrorValueRoundTripsThroughSaveAndAssertion(string variant)
    {
        using var fixture = new Fixture();
        var source = fixture.Source(variant);
        using var sessions = new ExcelSessions();
        var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
        var operations = new[]
        {
            JsonSerializer.Deserialize<SetValueRequest>("""{"op":"set_value","sheet":"Sheet1","target":"A1","rich_policy":"replace","value":{"error":"#N/A"}}""")!,
            JsonSerializer.Deserialize<SetValueRequest>("""{"op":"set_values","sheet":"Sheet1","target":"B1:C1","values":[[{"error":"#REF!"},{"error":"#DIV/0!"}]]}""")!,
            JsonSerializer.Deserialize<SetValueRequest>("""{"op":"set_value","sheet":"Sheet1","target":"E5","value":{"error":"#VALUE!"}}""")!
        };
        sessions.Apply(id, 0, operations.SelectMany(operation => operation.NormalizeMany(null)).ToArray());
        var output = Path.Combine(fixture.Directory, "errors.xlsx");
        var assertions = new[]
        {
            JsonSerializer.Deserialize<SaveAssertionRequest>("""{"target":"Sheet1!A1","equals":{"value":{"error":"#N/A"}}}""")!.Normalize(),
            JsonSerializer.Deserialize<SaveAssertionRequest>("""{"target":"Sheet1!C1","equals":{"value":{"error":"#DIV/0!"}}}""")!.Normalize()
        };
        var result = sessions.Save(id, output, assertions);
        Assert.Equal("verified", result.GetType().GetProperty("status")!.GetValue(result));
        var cells = P2aGates.ReadCells(output, "Sheet1", ["A1", "B1", "C1", "E5"]);
        Assert.Equal(["#N/A", "#REF!", "#DIV/0!", "#VALUE!"], cells.Select(cell => cell.Value));
        Assert.All(cells, cell => Assert.Equal("error", cell.Kind));
        using var archive = ZipFile.OpenRead(output);
        using var sheet = new StreamReader(archive.Entries.Single(entry => entry.FullName.EndsWith("sheet1.xml", StringComparison.OrdinalIgnoreCase)).Open());
        Assert.Contains("t=\"e\"", sheet.ReadToEnd());
        sessions.Close(id, true);
    }

    [Theory]
    [InlineData("#DIV/0!")]
    [InlineData("#N/A")]
    [InlineData("#NAME?")]
    [InlineData("#NULL!")]
    [InlineData("#NUM!")]
    [InlineData("#REF!")]
    [InlineData("#VALUE!")]
    [InlineData("#GETTING_DATA")]
    [InlineData("#SPILL!")]
    [InlineData("#CALC!")]
    [InlineData("#BLOCKED!")]
    [InlineData("#CONNECT!")]
    [InlineData("#EXTERNAL!")]
    [InlineData("#FIELD!")]
    [InlineData("#UNKNOWN!")]
    public void RecognizesSameErrorTokensAsLegacyFormulaCache(string token)
    {
        Assert.True(CellError.IsSupported(token));
    }

    [Theory]
    [InlineData("#NOT_A_REAL_ERROR!")]
    [InlineData("#REF")]
    [InlineData("#n/a")]
    [InlineData("")]
    public void RejectsUnsupportedErrorTokensBeforeMutating(string token)
    {
        using var fixture = new Fixture();
        using var store = new PackageStore(fixture.Source("default"));
        Assert.False(CellError.IsSupported(token));
        var request = JsonSerializer.Deserialize<SetValueRequest>(
            JsonSerializer.Serialize(new { op = "set_value", target = "B1", value = new { error = token } }))!;
        Assert.Throws<ArgumentException>(() => request.NormalizeMany("Sheet1"));
        Assert.Throws<FormatException>(() => SetValueEngine.Apply(store,
            [new SetValueOp("Sheet1", "B1", "error", token)]));
        Assert.Empty(store.ChangedParts);
        var assertion = JsonSerializer.Deserialize<SaveAssertionRequest>(
            JsonSerializer.Serialize(new { target = "Sheet1!B1", equals = new { value = new { error = token } } }))!;
        Assert.Throws<NotSupportedException>(() => assertion.Normalize());
    }

    [Fact]
    public void RejectsErrorObjectsWithExtraFields()
    {
        var request = JsonSerializer.Deserialize<SetValueRequest>(
            """{"op":"set_value","target":"B1","value":{"error":"#N/A","ignored":7}}""")!;
        Assert.Throws<ArgumentException>(() => request.NormalizeMany("Sheet1"));
        var assertion = JsonSerializer.Deserialize<SaveAssertionRequest>(
            """{"target":"Sheet1!B1","equals":{"value":{"error":"#N/A","ignored":7}}}""")!;
        Assert.Throws<NotSupportedException>(() => assertion.Normalize());
    }

    [Fact]
    public void LostErrorValueFailsG4AndIncorrectAssertionBlocksSave()
    {
        using var fixture = new Fixture();
        var source = fixture.Source("default");
        Assert.Contains(P2aGates.CheckIntent(source, [new CellExpectation("Sheet1", "B1", "error", "#N/A")]),
            issue => issue.Code == "INTENT_MISMATCH");
        using var sessions = new ExcelSessions();
        var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
        sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "error", "#N/A")]);
        var destination = Path.Combine(fixture.Directory, "blocked.xlsx");
        var assertion = JsonSerializer.Deserialize<SaveAssertionRequest>(
            """{"target":"Sheet1!B1","equals":{"value":{"error":"#REF!"}}}""")!.Normalize();
        var blocked = Assert.Throws<SaveBlockedException>(() => sessions.Save(id, destination, [assertion]));
        Assert.Contains(blocked.Issues, issue => issue.Gate == "G7" && issue.Code == "ASSERT_VALUE_MISMATCH");
        Assert.False(File.Exists(destination));
        sessions.Close(id, true);
    }

    private sealed class Fixture : IDisposable
    {
        public string Directory { get; } = Path.Combine(Path.GetTempPath(), "docloupe-errors-" + Guid.NewGuid().ToString("N"));
        public string Source(string variant)
        {
            SyntheticFixtures.Create(Directory);
            return Path.Combine(Directory, variant + ".xlsx");
        }
        public void Dispose() => System.IO.Directory.Delete(Directory, true);
    }
}
