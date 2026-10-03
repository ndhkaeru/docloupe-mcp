using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class FindTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void BoundedSearchUsesCurrentSessionAndNfcNormalization(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-find-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var initial = JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "A1:D3", "HELLO", false));
            Assert.True(initial.GetProperty("partial").GetBoolean());
            Assert.Equal(12, initial.GetProperty("total_scanned").GetInt32());
            Assert.Equal("Sheet1!A1", Assert.Single(initial.GetProperty("matches").EnumerateArray()).GetProperty("addr").GetString());
            Assert.False(initial.GetProperty("truncated").GetBoolean());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "A1:D3", "HELLO", false,
                caseSensitive: true)).GetProperty("matches").EnumerateArray());
            Assert.Equal("Sheet1!C1", Assert.Single(JsonSerializer.SerializeToElement(sessions.Find(id, null,
                "Sheet1!A1:D3", "1\\+1", true, searchIn: "formula")).GetProperty("matches").EnumerateArray())
                .GetProperty("addr").GetString());
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "E4", "text", "cafe\u0301")]);
            var normalized = JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "E4", "café", false));
            Assert.Equal(1, normalized.GetProperty("revision").GetInt32());
            Assert.Single(normalized.GetProperty("matches").EnumerateArray());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "E4", "café", false,
                normalize: "none")).GetProperty("matches").EnumerateArray());
            Assert.False(normalized.GetProperty("truncated").GetBoolean());
            sessions.Apply(id, 1, [new SetValueOp("Sheet1", "F4", "text", "cafe\u0301")]);
            var limited = JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "E4:F4", "café", false,
                maxResults: 1));
            Assert.Single(limited.GetProperty("matches").EnumerateArray());
            Assert.True(limited.GetProperty("truncated").GetBoolean());
            Assert.Equal(2, limited.GetProperty("total_scanned").GetInt32());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void SearchLimitsAndQueryShapeFailClosed()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-find-reject-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            foreach (var json in new[]
                {
                    "{}", "{\"text\":\"x\",\"regex\":\"y\"}",
                    "{\"text\":\"x\",\"style\":{}}", "{\"regex\":\"\"}"
                })
                Assert.ThrowsAny<Exception>(() => JsonSerializer.Deserialize<FindQueryRequest>(json)!.Normalize());
            Assert.ThrowsAny<Exception>(() => sessions.Find(id, "Sheet1", "A1:A501", "x", false));
            Assert.Throws<NotSupportedException>(() => sessions.Find(id, "Sheet1", "A1", "x", false, searchIn: "display"));
            Assert.ThrowsAny<Exception>(() => sessions.Find(id, "Sheet1", "A1", "[", true));
            Assert.ThrowsAny<Exception>(() => sessions.Find(id, "Sheet1", "A1", "x", false, maxResults: 101));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }
}
