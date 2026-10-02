using System.Text.Json;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ReaderFixtureTests
{
    [Theory]
    [InlineData("opc-percent-case")]
    [InlineData("nested-workbook")]
    public void SessionlessPeekProducesBoundedMarkdown(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-peek-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var path = Path.Combine(directory, variant + ".xlsx");
            var result = JsonSerializer.SerializeToElement(sessions.Peek(path, "preview", "Sheet1", 3, 4));
            Assert.True(result.GetProperty("partial").GetBoolean());
            Assert.Equal("Sheet1", result.GetProperty("sheets")[0].GetProperty("name").GetString());
            var preview = result.GetProperty("preview")[0];
            Assert.Equal("Sheet1", preview.GetProperty("sheet").GetString());
            var markdown = preview.GetProperty("markdown").GetString()!;
            Assert.Contains("| row | A | B | C | D |", markdown);
            Assert.Contains("| 1 | hello | 42 | 2 |", markdown);
            Assert.Contains("| 3 |", markdown);
            Assert.Contains("old |", markdown);
            var narrow = JsonSerializer.SerializeToElement(sessions.Peek(path, "preview", maxRows: 1, maxCols: 2))
                .GetProperty("preview")[0].GetProperty("markdown").GetString()!;
            Assert.Contains("| row | A | B |", narrow);
            Assert.DoesNotContain("| C |", narrow);
            Assert.DoesNotContain("| 2 |", narrow);
            Assert.Throws<KeyNotFoundException>(() => sessions.Peek(path, "preview", "Missing"));
            Assert.Throws<ArgumentOutOfRangeException>(() => sessions.Peek(path, "preview", maxRows: 101));
            Assert.Throws<ArgumentException>(() => sessions.Peek(path, "full"));
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Status()).GetProperty("sessions").EnumerateArray());
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void PreviewAndPartialVerificationReadEverySyntheticVariant(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-reader-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var path = Path.Combine(directory, variant + ".xlsx");
            var preview = WorkbookReader.Peek(path);
            Assert.Equal("Sheet1", Assert.Single(preview.Sheets).Name);
            var cells = preview.FirstSheetCells.ToDictionary(cell => cell.Address);
            Assert.Equal("hello", cells["A1"].Value);
            Assert.Equal("42", cells["B1"].Value);
            Assert.Equal("1+1", cells["C1"].Formula);
            Assert.Equal("2", cells["C1"].Value);
            Assert.Equal("old", cells["D3"].Value);
            Assert.Equal("unverified", WorkbookReader.VerifyPartial(path).Status);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }
}
