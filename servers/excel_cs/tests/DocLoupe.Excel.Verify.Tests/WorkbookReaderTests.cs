using System.IO.Compression;
using System.Text;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Verify.Tests;

public sealed class WorkbookReaderTests
{
    [Theory]
    [InlineData(false, "unverified")]
    [InlineData(true, "failed")]
    public void ReadsPrefixedContentTypesAndReportsOnlyImplementedGates(bool invalidMarkup, string expectedStatus)
    {
        var path = Path.Combine(AppContext.BaseDirectory, $"read-probe-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var archive = new ZipArchive(File.Create(path), ZipArchiveMode.Create))
            {
                Write(archive, "[Content_Types].xml", "<ns0:Types xmlns:ns0='http://schemas.openxmlformats.org/package/2006/content-types'/>");
                Write(archive, "xl/workbook.xml", "<x:workbook xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><x:sheets><x:sheet name='S' sheetId='1' r:id='rId1'/></x:sheets></x:workbook>");
                Write(archive, "xl/_rels/workbook.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'/></Relationships>");
                var ignorable = invalidMarkup ? "good missing" : "good";
                Write(archive, "xl/worksheets/sheet1.xml", $"<x:worksheet xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:good='urn:good' mc:Ignorable='{ignorable}'><x:sheetData><x:row r='1'><x:c r='A1' t='inlineStr'><x:is><x:t>Hello</x:t></x:is></x:c></x:row></x:sheetData></x:worksheet>");
            }
            var preview = WorkbookReader.Peek(path);
            Assert.Equal("S", Assert.Single(preview.Sheets).Name);
            Assert.Equal("Hello", Assert.Single(preview.FirstSheetCells).RawValue);
            var verified = WorkbookReader.VerifyPartial(path);
            Assert.Equal(expectedStatus, verified.Status);
            Assert.Contains("G5", verified.UnverifiedGates);
            if (invalidMarkup)
                Assert.Contains(verified.MarkupIssues, issue => issue.Code == "UNDECLARED_MC_PREFIX");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ExcludesPhoneticGuideTextFromCellValue()
    {
        // 漢字 with the phonetic guide かんじ: the value is the base text only.
        var path = CreateWorkbook("<x:row r='1'><x:c r='A1' t='inlineStr'><x:is><x:r><x:t>漢字</x:t></x:r><x:rPh sb='0' eb='2'><x:t>かんじ</x:t></x:rPh><x:phoneticPr fontId='1'/></x:is></x:c></x:row>");
        try
        {
            Assert.Equal("漢字", Assert.Single(WorkbookReader.Peek(path).FirstSheetCells).RawValue);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ReadsCachedValueThatDirectlyFollowsFormula()
    {
        var path = CreateWorkbook("<x:row r='1'><x:c r='A1'><x:v>10</x:v></x:c><x:c r='B1'><x:f>A1*2</x:f><x:v>20</x:v></x:c></x:row>");
        try
        {
            var cells = WorkbookReader.Peek(path).FirstSheetCells;
            Assert.Equal("10", cells[0].RawValue);
            Assert.Equal("A1*2", cells[1].Formula);
            Assert.Equal("20", cells[1].RawValue);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static string CreateWorkbook(string sheetDataContent)
    {
        var path = Path.Combine(AppContext.BaseDirectory, $"read-probe-{Guid.NewGuid():N}.xlsx");
        using var archive = new ZipArchive(File.Create(path), ZipArchiveMode.Create);
        Write(archive, "[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>");
        Write(archive, "xl/workbook.xml", "<x:workbook xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><x:sheets><x:sheet name='S' sheetId='1' r:id='rId1'/></x:sheets></x:workbook>");
        Write(archive, "xl/_rels/workbook.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'/></Relationships>");
        Write(archive, "xl/worksheets/sheet1.xml", $"<x:worksheet xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><x:sheetData>{sheetDataContent}</x:sheetData></x:worksheet>");
        return path;
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        using var stream = archive.CreateEntry(name).Open();
        stream.Write(Encoding.UTF8.GetBytes(content));
    }
}
