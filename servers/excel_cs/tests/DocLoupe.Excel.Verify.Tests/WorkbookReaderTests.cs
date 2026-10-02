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
                Write(archive, "_rels/.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId0' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/></Relationships>");
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
            var cell = Assert.Single(WorkbookReader.Peek(path).FirstSheetCells);
            Assert.Equal("漢字", cell.RawValue);
            Assert.Equal("漢字", cell.Value);
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
            Assert.Equal("20", cells[1].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("XL/Workbook.XML", "/xl/WORKBOOK.XML", "worksheets/sheet%31.xml", "XL/worksheets/sheet1.xml")]
    [InlineData("xl/nested/workbook.xml", "/xl/nested/workbook.xml", "/xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml")]
    [InlineData("xl/nested/workbook.xml", "xl/nested/workbook.xml", "../worksheets/sheet%31.xml", "xl/worksheets/sheet1.xml")]
    [InlineData("xl/nested/workbook.xml", "xl/nested/work%62ook.xml", "/xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml")]
    public void ResolvesWorkbookAndSheetThroughOpcRelationships(string workbookPart, string workbookTarget,
        string sheetTarget, string sheetPart)
    {
        var path = CreateWorkbook("<x:row r='1'><x:c r='A1'><x:v>42</x:v></x:c></x:row>",
            workbookPart, workbookTarget, sheetTarget, sheetPart);
        try
        {
            var preview = WorkbookReader.Peek(path);
            Assert.Equal(sheetPart, Assert.Single(preview.Sheets).Part);
            Assert.Equal("42", Assert.Single(preview.FirstSheetCells).RawValue);
            Assert.Equal("unverified", WorkbookReader.VerifyPartial(path).Status);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(null, false, "MISSING_PART")]
    [InlineData("/../xl/workbook.xml", true, "INVALID_ROOT_RELATIONSHIPS")]
    [InlineData("/xl/missing.xml", true, "MISSING_PART")]
    [InlineData("/%2E%2E/xl/workbook.xml", true, "INVALID_ROOT_RELATIONSHIPS")]
    public void RejectsMissingOrInvalidMainRelationship(string? target, bool includeRoot, string issueCode)
    {
        var path = CreateWorkbook("", workbookTarget: target, includeRoot: includeRoot);
        try
        {
            Assert.Throws<InvalidDataException>(() => WorkbookReader.Peek(path));
            var result = WorkbookReader.VerifyPartial(path);
            Assert.Equal("failed", result.Status);
            Assert.Contains(result.PackageIssues, issue => issue.Code == issueCode);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DoesNotTreatAnUnrelatedRootRelationshipAsAWorkbook()
    {
        var path = CreateWorkbook("", rootRelationshipType: "worksheet");
        try
        {
            Assert.Throws<InvalidDataException>(() => WorkbookReader.Peek(path));
            Assert.Contains(WorkbookReader.VerifyPartial(path).PackageIssues,
                issue => issue.Code == "INVALID_ROOT_RELATIONSHIPS");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("worksheets/missing.xml", "worksheet", "MISSING_PART")]
    [InlineData("../../../outside.xml", "worksheet", "INVALID_WORKBOOK_STRUCTURE")]
    [InlineData("worksheets/sheet1.xml", "externalLink", "INVALID_WORKBOOK_STRUCTURE")]
    public void ReportsMissingOrInvalidWorksheetRelationships(string sheetTarget, string relationshipType, string issueCode)
    {
        var path = CreateWorkbook("", sheetTarget: sheetTarget, sheetRelationshipType: relationshipType);
        try
        {
            Assert.Throws<InvalidDataException>(() => WorkbookReader.Peek(path));
            var result = WorkbookReader.VerifyPartial(path);
            Assert.Equal("failed", result.Status);
            Assert.Contains(result.PackageIssues, issue => issue.Code == issueCode);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SelectsRequestedSheetAndBoundsThePreview()
    {
        var path = CreateWorkbook("<x:row r='1'><x:c r='A1'><x:v>first</x:v></x:c></x:row>",
            secondSheetContent: "<x:row r='1'><x:c r='A1'><x:v>second</x:v></x:c><x:c r='B1'><x:v>outside</x:v></x:c></x:row>");
        try
        {
            var workbook = WorkbookReader.Peek(path, maxCells: 2, sheetName: "Second", maxRows: 1, maxColumns: 1);
            Assert.Equal(2, workbook.Sheets.Count);
            Assert.Equal("second", Assert.Single(workbook.FirstSheetCells).Value);
            Assert.Throws<KeyNotFoundException>(() => WorkbookReader.Peek(path, sheetName: "Missing"));
            Assert.Throws<ArgumentOutOfRangeException>(() => WorkbookReader.Peek(path, maxRows: 0));
            Assert.Equal("unverified", WorkbookReader.VerifyPartial(path).Status);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ComputesUsedRangeFromCellsRatherThanStaleDimension()
    {
        var path = CreateWorkbook("<x:row r='4'><x:c r='D4'><x:v>1</x:v></x:c></x:row>"
            + "<x:row r='8'><x:c r='G8'><x:v>2</x:v></x:c></x:row>"
            + "<x:row r='20' s='2' customFormat='1'/>",
            secondSheetContent: "", worksheetDimension: "A1:Z999",
            worksheetTrailing: "<x:extLst><x:ext><x:c r='Z99'><x:v>ghost</x:v></x:c></x:ext></x:extLst>");
        try
        {
            var sheets = WorkbookReader.Peek(path, maxCells: 0).Sheets;
            Assert.Equal("D4:G8", sheets[0].UsedRange);
            Assert.Null(sheets[1].UsedRange);
            Assert.Equal(new[] { "D4", "G8" }, WorkbookReader.Peek(path).FirstSheetCells.Select(cell => cell.Address));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<x:row r='2'><x:c r='A3'><x:v>1</x:v></x:c></x:row>")]
    [InlineData("<x:row r='0'><x:c r='A1'><x:v>1</x:v></x:c></x:row>")]
    [InlineData("<x:row r='1'><x:c r='XFE1'><x:v>1</x:v></x:c></x:row>")]
    public void RejectsInvalidCellCoordinatesWhenComputingUsedRange(string row)
    {
        var path = CreateWorkbook(row);
        try
        {
            Assert.Throws<InvalidDataException>(() => WorkbookReader.Peek(path, maxCells: 0));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("0", "漢字")]
    [InlineData("1", "Other")]
    public void ResolvesSharedStringsWithoutIncludingPhoneticGuide(string index, string expected)
    {
        var strings = "<x:sst xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>"
            + "<x:si><x:r><x:t>漢字</x:t></x:r><x:rPh sb='0' eb='2'><x:t>かんじ</x:t></x:rPh></x:si>"
            + "<x:si><x:t>Other</x:t></x:si></x:sst>";
        var path = CreateWorkbook($"<x:row r='1'><x:c r='A1' t='s'><x:v>{index}</x:v></x:c></x:row>",
            workbookPart: "xl/nested/workbook.xml", sheetTarget: "/xl/worksheets/sheet1.xml", sharedStringsXml: strings);
        try
        {
            var cell = Assert.Single(WorkbookReader.Peek(path).FirstSheetCells);
            Assert.Equal(index, cell.RawValue);
            Assert.Equal(expected, cell.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("2")]
    [InlineData("-1")]
    public void RejectsInvalidSharedStringIndices(string index)
    {
        var strings = "<x:sst xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><x:si><x:t>One</x:t></x:si></x:sst>";
        var path = CreateWorkbook($"<x:row r='1'><x:c r='A1' t='s'><x:v>{index}</x:v></x:c></x:row>",
            sharedStringsXml: strings);
        try
        {
            Assert.Throws<InvalidDataException>(() => WorkbookReader.Peek(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("../work%62ook.xml", null, null)]
    [InlineData("../workbook.xml#fragment", "Internal", null)]
    [InlineData("https://example.invalid/resource", "External", null)]
    [InlineData("../missing.xml", null, "MISSING_RELATIONSHIP_TARGET")]
    [InlineData("../../../escape.xml", null, "INVALID_RELATIONSHIP_TARGET")]
    [InlineData("../workbook.xml", "Unknown", "INVALID_TARGET_MODE")]
    public void ChecksUnmodifiedWorksheetRelationships(string target, string? mode, string? expectedCode)
    {
        var path = CreateWorkbook("");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
                Write(archive, "xl/worksheets/_rels/sheet1.xml.rels",
                    $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
                    $"<Relationship Id='rId1' Type='urn:custom' Target='{target}'" +
                    (mode is null ? "" : $" TargetMode='{mode}'") + "/></Relationships>");
            var result = WorkbookReader.VerifyPartial(path);
            if (expectedCode is null)
                Assert.Equal("unverified", result.Status);
            else
            {
                Assert.Equal("failed", result.Status);
                Assert.Contains(result.PackageIssues, issue => issue.Code == expectedCode);
            }
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void RejectsDuplicateIdsAndOrphanRelationshipSources()
    {
        var path = CreateWorkbook("");
        try
        {
            const string relationship = "<Relationship Id='dup' Type='urn:custom' Target='../workbook.xml'/>";
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                Write(archive, "xl/worksheets/_rels/sheet1.xml.rels",
                    "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
                    relationship + relationship + "</Relationships>");
                Write(archive, "xl/worksheets/_rels/ghost.xml.rels",
                    "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'/>");
            }
            var result = WorkbookReader.VerifyPartial(path);
            Assert.Equal("failed", result.Status);
            Assert.Contains(result.PackageIssues, issue => issue.Code == "DUPLICATE_RELATIONSHIP_ID");
            Assert.Contains(result.PackageIssues, issue => issue.Code == "MISSING_SOURCE_PART");
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false, false, "UNRESOLVED_RELATIONSHIP_ID")]
    [InlineData(true, false, "MISSING_RELATIONSHIP_TARGET")]
    [InlineData(true, true, null)]
    public void ResolvesRelationshipIdsInUnmodifiedXmlParts(bool includeRelationship, bool includeTarget, string? expectedCode)
    {
        var path = CreateWorkbook("", worksheetTrailing:
            "<x:drawing xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' r:id='rId8'/>");
        try
        {
            if (includeRelationship)
            {
                using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
                Write(archive, "xl/worksheets/_rels/sheet1.xml.rels",
                    "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
                    "<Relationship Id='rId8' Type='urn:drawing' Target='../drawings/drawing1.xml'/>" +
                    "</Relationships>");
                if (includeTarget)
                    Write(archive, "xl/drawings/drawing1.xml", "<drawing xmlns='urn:drawing'/>");
            }
            var result = WorkbookReader.VerifyPartial(path);
            if (expectedCode is null)
                Assert.Equal("unverified", result.Status);
            else
            {
                Assert.Equal("failed", result.Status);
                Assert.Contains(result.PackageIssues, issue => issue.Code == expectedCode);
            }
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("embed")]
    [InlineData("link")]
    [InlineData("pict")]
    public void ResolvesOtherNamespacedRelationshipReferences(string attribute)
    {
        var path = CreateWorkbook("", worksheetTrailing:
            $"<x:drawing xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' r:{attribute}='rId8'/>");
        try
        {
            var missing = WorkbookReader.VerifyPartial(path);
            Assert.Equal("failed", missing.Status);
            Assert.Contains(missing.PackageIssues, issue => issue.Code == "UNRESOLVED_RELATIONSHIP_ID"
                && issue.Detail.Contains($"@{attribute}=rId8", StringComparison.Ordinal));
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
                Write(archive, "xl/worksheets/_rels/sheet1.xml.rels",
                    "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
                    "<Relationship Id='rId8' Type='urn:custom' Target='../workbook.xml'/></Relationships>");
            Assert.Equal("unverified", WorkbookReader.VerifyPartial(path).Status);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DoesNotTreatOtherIdNamespacesAsRelationships()
    {
        var path = CreateWorkbook("", worksheetTrailing: "<x:drawing xmlns:custom='urn:custom' custom:id='not-a-rel'/>");
        try { Assert.Equal("unverified", WorkbookReader.VerifyPartial(path).Status); }
        finally { File.Delete(path); }
    }

    private static string CreateWorkbook(string sheetDataContent, string workbookPart = "xl/workbook.xml",
        string? workbookTarget = null, string sheetTarget = "worksheets/sheet1.xml",
        string sheetPart = "xl/worksheets/sheet1.xml", bool includeRoot = true,
        string rootRelationshipType = "officeDocument", string sheetRelationshipType = "worksheet",
        string? sharedStringsXml = null, string? secondSheetContent = null, string? worksheetDimension = null,
        string worksheetTrailing = "")
    {
        var path = Path.Combine(AppContext.BaseDirectory, $"read-probe-{Guid.NewGuid():N}.xlsx");
        using var archive = new ZipArchive(File.Create(path), ZipArchiveMode.Create);
        Write(archive, "[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>");
        if (includeRoot)
            Write(archive, "_rels/.rels", $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId0' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rootRelationshipType}' Target='{workbookTarget ?? workbookPart}'/></Relationships>");
        var secondSheet = secondSheetContent is null ? "" : "<x:sheet name='Second' sheetId='2' r:id='rId3'/>";
        Write(archive, workbookPart, $"<x:workbook xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><x:sheets><x:sheet name='S' sheetId='1' r:id='rId1'/>{secondSheet}</x:sheets></x:workbook>");
        var slash = workbookPart.LastIndexOf('/');
        var relationshipPart = (slash < 0 ? "" : workbookPart[..(slash + 1)]) + "_rels/" + workbookPart[(slash + 1)..] + ".rels";
        var sharedRelationship = sharedStringsXml is null ? "" : "<Relationship Id='rId2' Target='sharedStrings.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'/>";
        var secondRelationship = secondSheetContent is null ? "" : "<Relationship Id='rId3' Target='worksheets/sheet2.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'/>";
        Write(archive, relationshipPart, $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='{sheetTarget}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/{sheetRelationshipType}'/>{sharedRelationship}{secondRelationship}</Relationships>");
        if (sharedStringsXml is not null)
            Write(archive, (slash < 0 ? "" : workbookPart[..(slash + 1)]) + "sharedStrings.xml", sharedStringsXml);
        var dimension = worksheetDimension is null ? "" : $"<x:dimension ref='{worksheetDimension}'/>";
        Write(archive, sheetPart, $"<x:worksheet xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>{dimension}<x:sheetData>{sheetDataContent}</x:sheetData>{worksheetTrailing}</x:worksheet>");
        if (secondSheetContent is not null)
        {
            var secondPart = (slash < 0 ? "" : workbookPart[..(slash + 1)]) + "worksheets/sheet2.xml";
            Write(archive, secondPart, $"<x:worksheet xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><x:sheetData>{secondSheetContent}</x:sheetData></x:worksheet>");
        }
        return path;
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        using var stream = archive.CreateEntry(name).Open();
        stream.Write(Encoding.UTF8.GetBytes(content));
    }
}
