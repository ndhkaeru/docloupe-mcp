using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ReaderFixtureTests
{
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
