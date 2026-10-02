using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ReadRangeTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void ReadsBoundedRectangleFromCurrentRevision(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-range-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var initial = Read(sessions, session, "Sheet1!A1:C2", "WrongDefault");
            Assert.Equal(0, initial.GetProperty("revision").GetInt32());
            Assert.Equal("hello", Value(initial, "A1"));
            Assert.Equal("42", Value(initial, "B1"));
            Assert.Equal("1+1", Formula(initial, "C1"));
            Assert.Equal(3, initial.GetProperty("cells").GetArrayLength());
            sessions.Apply(session, 0, [new SetValueOp("Sheet1", "B2", "number", "27")]);
            var preview = Read(sessions, session, "B1:C2", "Sheet1");
            Assert.Equal(1, preview.GetProperty("revision").GetInt32());
            Assert.Equal("27", Value(preview, "B2"));
            Assert.DoesNotContain(preview.GetProperty("cells").EnumerateArray(),
                cell => cell.GetProperty("Address").GetString() == "A1");
            sessions.Close(session, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("B2:A2")]
    [InlineData("B1:A2")]
    [InlineData("A1:B1:C1")]
    [InlineData("A1:Sheet1!B1")]
    [InlineData("A1:A501")]
    [InlineData("XFD1048576:XFD1048577")]
    public void RejectsInvalidOrTooLargeRanges(string target)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-reject-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.ThrowsAny<Exception>(() => sessions.Read(session, "Sheet1", [target]));
            Assert.Equal("hello", Value(Read(sessions, session, "A1:A500", "Sheet1"), "A1"));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(session)).GetProperty("revision").GetInt32());
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void RejectsMixedSheetTargetsWithoutReadingAnyWorkbook()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-mixed-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.Throws<ArgumentException>(() => sessions.Read(session, "Sheet1", ["A1:B1", "Other!A1"]));
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static JsonElement Read(ExcelSessions sessions, string session, string target, string sheet) =>
        JsonSerializer.SerializeToElement(sessions.Read(session, sheet, [target]));

    private static string? Value(JsonElement result, string address) => result.GetProperty("cells").EnumerateArray()
        .Single(cell => cell.GetProperty("Address").GetString() == address).GetProperty("Value").GetString();

    private static string? Formula(JsonElement result, string address) => result.GetProperty("cells").EnumerateArray()
        .Single(cell => cell.GetProperty("Address").GetString() == address).GetProperty("Formula").GetString();
}
