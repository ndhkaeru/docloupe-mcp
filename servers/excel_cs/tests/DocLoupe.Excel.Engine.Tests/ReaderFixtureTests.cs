using System.IO.Compression;
using System.Text.Json;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ReaderFixtureTests
{
    [Fact]
    public void SessionlessVerifyReportsPartialSuccessAndCorruptionWithoutChangingFiles()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-verify-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var source = Path.Combine(directory, "default.xlsx");
            var sourceBytes = File.ReadAllBytes(source);
            var valid = sessions.Verify(source);
            Assert.Equal(Path.GetFullPath(source), valid.Path);
            Assert.Equal("unverified", valid.Summary.Status);
            Assert.Contains("G2", valid.Summary.UnverifiedGates);
            Assert.Equal(sourceBytes, File.ReadAllBytes(source));

            var broken = Path.Combine(directory, "broken.xlsx");
            File.Copy(source, broken);
            using (var archive = ZipFile.Open(broken, ZipArchiveMode.Update))
                archive.GetEntry("_rels/.rels")!.Delete();
            var brokenBytes = File.ReadAllBytes(broken);
            var invalid = sessions.Verify(broken);
            Assert.Equal("failed", invalid.Summary.Status);
            Assert.Contains(invalid.Summary.PackageIssues, issue => issue.Code == "MISSING_PART");
            Assert.Equal(brokenBytes, File.ReadAllBytes(broken));
            var invalidZip = Path.Combine(directory, "invalid-zip.xlsx");
            File.WriteAllText(invalidZip, "not a zip archive");
            var zipResult = sessions.Verify(invalidZip);
            Assert.Equal("failed", zipResult.Summary.Status);
            Assert.Contains(zipResult.Summary.PackageIssues, issue => issue.Code == "INVALID_PACKAGE");
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Status()).GetProperty("sessions").EnumerateArray());
            Assert.Throws<FileNotFoundException>(() => sessions.Verify(Path.Combine(directory, "missing.xlsx")));
            var unsupported = Path.Combine(directory, "default.txt");
            File.Copy(source, unsupported);
            Assert.Throws<NotSupportedException>(() => sessions.Verify(unsupported));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }

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
            Assert.Equal("A1:D3", result.GetProperty("sheets")[0].GetProperty("used_range").GetString());
            Assert.Equal("explicit_cells", result.GetProperty("used_range_basis").GetString());
            Assert.DoesNotContain("used_range", result.GetProperty("unverified").EnumerateArray()
                .Select(item => item.GetString()));
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
            var selected = Assert.Single(preview.Sheets);
            Assert.Equal("Sheet1", selected.Name);
            Assert.Equal("A1:D3", selected.UsedRange);
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
