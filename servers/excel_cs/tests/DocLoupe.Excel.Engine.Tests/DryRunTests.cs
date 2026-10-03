using System.Text.Json;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class DryRunTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void PlanMatchesApplyWithoutChangingSessionOrSource(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-dry-run-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, variant + ".xlsx");
            var original = File.ReadAllBytes(source);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("""{"op":"set_value","target":"Sheet1!F4","value":17}""")]);
            var operations = new[]
            {
                Request("""{"op":"set_value","target":"Sheet1!B1","value":"replacement","expect":{"value":42}}"""),
                Request("""{"op":"set_value","target":"Sheet1!G4","value":true,"expect":{"value":null}}""")
            };

            var plan = JsonSerializer.SerializeToElement(sessions.Apply(id, 1, operations, dryRun: true));
            Assert.True(plan.GetProperty("dry_run").GetBoolean());
            Assert.Equal(1, plan.GetProperty("revision_before").GetInt32());
            Assert.Equal(1, plan.GetProperty("revision_after").GetInt32());
            Assert.Equal(1, plan.GetProperty("revision").GetInt32());
            var plannedCells = plan.GetProperty("readback");
            Assert.Equal(2, plannedCells.EnumerateObject().Count());
            Assert.Equal("replacement", plannedCells.GetProperty("Sheet1!B1").GetProperty("Value").GetString());
            Assert.Equal("true", plannedCells.GetProperty("Sheet1!G4").GetProperty("Value").GetString());
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("ledger").GetArrayLength());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["G4"]))
                .GetProperty("cells").EnumerateArray());
            Assert.Equal(original, File.ReadAllBytes(source));

            var applied = JsonSerializer.SerializeToElement(sessions.Apply(id, 1, operations));
            Assert.False(applied.GetProperty("dry_run").GetBoolean());
            Assert.Equal(2, applied.GetProperty("revision_after").GetInt32());
            Assert.Equal(plan.GetProperty("intent").GetRawText(), applied.GetProperty("intent").GetRawText());
            Assert.Equal(plan.GetProperty("changed_parts").GetRawText(), applied.GetProperty("changed_parts").GetRawText());
            Assert.Equal(plannedCells.GetRawText(), applied.GetProperty("readback").GetRawText());
            var output = Path.Combine(directory, "saved.xlsx");
            sessions.Save(id, output);
            Assert.Equal("replacement", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            Assert.Equal("true", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["G4"])).Value);
            Assert.Equal(original, File.ReadAllBytes(source));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ReadbackIncludesRemovedCellAsNull()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-dry-clear-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            var request = Request("""{"op":"clear","target":"Sheet1!B1","remove_cells":true}""");
            var plan = JsonSerializer.SerializeToElement(sessions.Apply(id, 0, [request], dryRun: true));
            Assert.Equal(JsonValueKind.Null, plan.GetProperty("readback").GetProperty("Sheet1!B1").ValueKind);
            var applied = JsonSerializer.SerializeToElement(sessions.Apply(id, 0, [request]));
            Assert.Equal(JsonValueKind.Null, applied.GetProperty("readback").GetProperty("Sheet1!B1").ValueKind);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void PreconditionRevisionAndWriterFailuresDoNotChangeSession()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-dry-failure-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.Throws<InvalidOperationException>(() => sessions.Apply(id, 1,
                [Request("""{"op":"set_value","target":"Sheet1!B1","value":3}""")], dryRun: true));
            Assert.Throws<PreconditionFailedException>(() => sessions.Apply(id, 0,
                [Request("""{"op":"set_value","target":"Sheet1!B1","value":3,"expect":{"value":0}}""")], dryRun: true));
            Assert.ThrowsAny<Exception>(() => sessions.Apply(id, 0,
                [Request("""{"op":"set_value","target":"MissingSheet!A1","value":3}""")], dryRun: true));
            var status = JsonSerializer.SerializeToElement(sessions.Status(id));
            Assert.Equal(0, status.GetProperty("revision").GetInt32());
            Assert.Empty(status.GetProperty("ledger").EnumerateArray());
            Assert.Throws<InvalidOperationException>(() => sessions.Save(id, Path.Combine(directory, "empty.xlsx")));
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static DocLoupe.Excel.Engine.SetValueOp Request(string json) =>
        Assert.Single(JsonSerializer.Deserialize<SetValueRequest>(json)!.NormalizeMany(null));
}
