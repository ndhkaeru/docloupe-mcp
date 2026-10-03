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
            var plannedResults = plan.GetProperty("results");
            Assert.Equal(2, plannedResults.GetArrayLength());
            Assert.Equal(0, plannedResults[0].GetProperty("index").GetInt32());
            Assert.Equal("set_value", plannedResults[0].GetProperty("op").GetString());
            Assert.Equal("planned", plannedResults[0].GetProperty("status").GetString());
            Assert.Equal("Sheet1!B1", plannedResults[0].GetProperty("resolved").GetString());
            Assert.Equal(1, plannedResults[1].GetProperty("index").GetInt32());
            Assert.Equal("Sheet1!G4", plannedResults[1].GetProperty("resolved").GetString());
            var plannedCells = plan.GetProperty("readback");
            Assert.Equal(2, plannedCells.EnumerateObject().Count());
            Assert.Equal("replacement", plannedCells.GetProperty("Sheet1!B1").GetProperty("Value").GetString());
            Assert.Equal("true", plannedCells.GetProperty("Sheet1!G4").GetProperty("Value").GetString());
            var differences = plan.GetProperty("diff").EnumerateArray().ToArray();
            var valueChange = Assert.Single(differences, item =>
                item.GetProperty("path").GetString() == "Sheet1!B1.value");
            Assert.Equal("42", valueChange.GetProperty("before").GetString());
            Assert.Equal("replacement", valueChange.GetProperty("after").GetString());
            Assert.StartsWith("d_", valueChange.GetProperty("id").GetString());
            Assert.DoesNotContain(differences, item => item.GetProperty("path").GetString()!.Contains("F4"));
            Assert.True(plan.GetProperty("diff_summary").GetProperty("partial").GetBoolean());
            Assert.False(plan.GetProperty("diff_summary").GetProperty("truncated").GetBoolean());
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
            Assert.Equal(plan.GetProperty("diff").GetRawText(), applied.GetProperty("diff").GetRawText());
            Assert.Equal("applied", applied.GetProperty("results")[0].GetProperty("status").GetString());
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
            var removed = Assert.Single(plan.GetProperty("diff").EnumerateArray(), item =>
                item.GetProperty("path").GetString() == "Sheet1!B1.present");
            Assert.True(removed.GetProperty("before").GetBoolean());
            Assert.False(removed.GetProperty("after").GetBoolean());
            var applied = JsonSerializer.SerializeToElement(sessions.Apply(id, 0, [request]));
            Assert.Equal(JsonValueKind.Null, applied.GetProperty("readback").GetProperty("Sheet1!B1").ValueKind);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void DiffCapReportsTotalChangesWithoutSilentlyDroppingThem()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-dry-diff-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var operations = new[]
            {
                Request("""{"op":"set_value","target":"Sheet1!B1","value":"changed"}"""),
                Request("""{"op":"set_value","target":"Sheet1!F4","value":3}""")
            };
            var plan = JsonSerializer.SerializeToElement(sessions.Apply(id, 0, operations,
                dryRun: true, maxDiffItems: 1));
            Assert.Equal(1, plan.GetProperty("diff").GetArrayLength());
            var summary = plan.GetProperty("diff_summary");
            Assert.True(summary.GetProperty("truncated").GetBoolean());
            Assert.True(summary.GetProperty("facets_changed").GetInt32() > 1);
            Assert.Equal(2, summary.GetProperty("cells_touched").GetInt32());
            Assert.Equal(2, plan.GetProperty("readback").EnumerateObject().Count());
            Assert.Throws<ArgumentOutOfRangeException>(() => sessions.Apply(id, 0, operations,
                dryRun: true, maxDiffItems: 0));
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ResultsRetainOriginalIndexForExpandedRange()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-dry-results-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var expanded = JsonSerializer.Deserialize<SetValueRequest>(
                """{"op":"set_value","label":"seed row","target":"Sheet1!E6:F6","value":1}""")!
                .NormalizeMany(null).Select(cell => cell with { SourceIndex = 0 }).ToArray();
            var unlabeled = Request("""{"op":"set_value","target":"Sheet1!G6","value":2}""") with { SourceIndex = 1 };
            var plan = JsonSerializer.SerializeToElement(sessions.Apply(id, 0, [.. expanded, unlabeled], dryRun: true));
            var results = plan.GetProperty("results").EnumerateArray().ToArray();
            Assert.Equal(3, results.Length);
            Assert.All(results[..2], item => Assert.Equal(0, item.GetProperty("index").GetInt32()));
            Assert.All(results[..2], item => Assert.Equal("seed row", item.GetProperty("label").GetString()));
            Assert.Equal("Sheet1!E6", results[0].GetProperty("resolved").GetString());
            Assert.Equal("Sheet1!F6", results[1].GetProperty("resolved").GetString());
            Assert.Equal(1, results[2].GetProperty("index").GetInt32());
            Assert.False(results[2].TryGetProperty("label", out _));
            Assert.Equal(3, plan.GetProperty("readback").EnumerateObject().Count());
            var applied = JsonSerializer.SerializeToElement(sessions.Apply(id, 0, [.. expanded, unlabeled]));
            Assert.Equal("seed row", applied.GetProperty("results")[0].GetProperty("label").GetString());
            Assert.False(applied.GetProperty("results")[2].TryGetProperty("label", out _));
            var tooLong = JsonSerializer.Deserialize<SetValueRequest>(JsonSerializer.Serialize(new
            {
                op = "set_value", label = new string('x', 257), target = "Sheet1!B1", value = 1
            }))!;
            Assert.Throws<ArgumentException>(() => tooLong.NormalizeMany(null));
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
