using System.IO.Compression;
using System.Text;
using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class AdvancedPartGateTests
{
    private const string Office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Theory]
    [InlineData("xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")]
    [InlineData("xl/media/image1.png", "image/png")]
    [InlineData("xl/printerSettings/printerSettings1.bin", "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings")]
    [InlineData("xl/pivotCache/pivotCacheDefinition1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml")]
    [InlineData("xl/pivotTables/pivotTable1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml")]
    [InlineData("xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml")]
    [InlineData("xl/slicers/slicer1.xml", "application/xml")]
    [InlineData("xl/timelines/timeline1.xml", "application/xml")]
    [InlineData("xl/persons/person1.xml", "application/xml")]
    [InlineData("xl/externalLinks/externalLink1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml")]
    [InlineData("customXml/item1.xml", "application/xml")]
    [InlineData("customUI/customUI.xml", "application/xml")]
    [InlineData("xl/threadedComments/threadedComment1.xml", "application/xml")]
    [InlineData("xl/model/model.bin", "application/octet-stream")]
    [InlineData("xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")]
    public void ProtectedPartsCannotBeDeletedOrChanged(string part, string type)
    {
        using var fixture = new Fixture();
        using (var zip = ZipFile.Open(fixture.Source, ZipArchiveMode.Update))
        {
            Add(zip, part, "original");
            Modify(zip, "[Content_Types].xml", xml => xml.Replace("</Types>",
                $"<Override PartName=\"/{part}\" ContentType=\"{type}\"/></Types>", StringComparison.Ordinal));
        }
        var changed = fixture.Output("changed");
        File.Copy(fixture.Source, changed);
        using (var zip = ZipFile.Open(changed, ZipArchiveMode.Update)) Replace(zip, part, "damaged");
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, changed, []),
            issue => issue.Gate == "G6" && issue.Code == "PROTECTED_PART_CHANGED" && issue.Detail == part);

        var missing = fixture.Output("missing");
        File.Copy(fixture.Source, missing);
        using (var zip = ZipFile.Open(missing, ZipArchiveMode.Update)) zip.GetEntry(part)!.Delete();
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, missing, []),
            issue => issue.Gate == "G6" && issue.Code == "PROTECTED_PART_MISSING" && issue.Detail == part);
    }

    [Fact]
    public void VbaRelationshipAndContentTypeCannotDisappear()
    {
        using var fixture = new Fixture();
        using (var zip = ZipFile.Open(fixture.Source, ZipArchiveMode.Update))
        {
            Add(zip, "xl/vbaProject.bin", "macro");
            Modify(zip, "xl/_rels/workbook.xml.rels", xml => xml.Replace("</Relationships>",
                $"<Relationship Id=\"rId9\" Type=\"{Office}/vbaProject\" Target=\"vbaProject.bin\"/></Relationships>", StringComparison.Ordinal));
            Modify(zip, "[Content_Types].xml", xml => xml.Replace("</Types>",
                "<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/></Types>", StringComparison.Ordinal));
        }
        var unchanged = fixture.Output("unchanged");
        File.Copy(fixture.Source, unchanged);
        Assert.Empty(AdvancedPartGate.Check(fixture.Source, unchanged, []));

        var vbaChanged = fixture.Output("vba-changed");
        File.Copy(fixture.Source, vbaChanged);
        using (var zip = ZipFile.Open(vbaChanged, ZipArchiveMode.Update)) Replace(zip, "xl/vbaProject.bin", "different");
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, vbaChanged, []),
            issue => issue.Gate == "G6" && issue.Code == "PROTECTED_PART_CHANGED" && issue.Detail == "xl/vbaProject.bin");

        var relLost = fixture.Output("rel-lost");
        File.Copy(fixture.Source, relLost);
        using (var zip = ZipFile.Open(relLost, ZipArchiveMode.Update))
            Modify(zip, "xl/_rels/workbook.xml.rels", xml => xml.Replace(
                $"<Relationship Id=\"rId9\" Type=\"{Office}/vbaProject\" Target=\"vbaProject.bin\"/>", "", StringComparison.Ordinal));
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, relLost, []),
            issue => issue.Gate == "G6" && issue.Code == "RELATIONSHIP_CHANGED");

        var typeLost = fixture.Output("type-lost");
        File.Copy(fixture.Source, typeLost);
        using (var zip = ZipFile.Open(typeLost, ZipArchiveMode.Update))
            Modify(zip, "[Content_Types].xml", xml => xml.Replace("application/vnd.ms-office.vbaProject", "application/octet-stream", StringComparison.Ordinal));
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, typeLost, []),
            issue => issue.Gate == "G6" && issue.Code == "VBA_CONTENT_TYPE");
    }

    [Fact]
    public void WorksheetContentTypeCannotHideAdvancedPartMutation()
    {
        using var fixture = new Fixture();
        using (var zip = ZipFile.Open(fixture.Source, ZipArchiveMode.Update))
        {
            Add(zip, "xl/charts/chart1.xml", "chart");
            Modify(zip, "[Content_Types].xml", xml => xml.Replace("</Types>",
                "<Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>", StringComparison.Ordinal));
        }
        var output = fixture.Output("disguised");
        File.Copy(fixture.Source, output);
        using (var zip = ZipFile.Open(output, ZipArchiveMode.Update)) Replace(zip, "xl/charts/chart1.xml", "damaged");
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, output, []),
            issue => issue.Gate == "G6" && issue.Code == "PROTECTED_PART_CHANGED");
    }

    [Fact]
    public void SignedSourceBlocksSaveBeforeDestinationIsCreated()
    {
        using var fixture = new Fixture();
        using (var zip = ZipFile.Open(fixture.Source, ZipArchiveMode.Update)) Add(zip, "_xmlsignatures/sig1.xml", "signature");
        using var sessions = new ExcelSessions();
        var id = JsonSerializer.SerializeToElement(sessions.Open(fixture.Source)).GetProperty("session").GetString()!;
        sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "9")]);
        var output = fixture.Output("signed");
        var blocked = Assert.Throws<SaveBlockedException>(() => sessions.Save(id, output));
        Assert.Contains(blocked.Issues, issue => issue.Gate == "G6" && issue.Code == "SIGNED_PACKAGE_UNSUPPORTED");
        Assert.False(File.Exists(output));
        sessions.Close(id, true);
    }

    [Fact]
    public void CalculationChainMayOnlyBeRemovedWhenFormulaIsOverwritten()
    {
        using var fixture = new Fixture();
        AddCalculationChain(fixture.Source);
        foreach (var (address, kind, expectsRemoval) in new[]
            { ("B1", "number", false), ("C1", "formula", false), ("C1", "number", true) })
        {
            var output = fixture.Output(address + kind);
            using (var store = new PackageStore(fixture.Source))
            {
                SetValueEngine.Apply(store, [new SetValueOp("Sheet1", address, kind, "9")]);
                store.Save(output);
            }
            Assert.Empty(AdvancedPartGate.Check(fixture.Source, output,
                [new CellExpectation("Sheet1", address, kind, "9")]));
            using var zip = ZipFile.OpenRead(output);
            Assert.Equal(expectsRemoval, zip.GetEntry("xl/calcChain.xml") is null);
        }

        var lost = fixture.Output("chain-lost");
        File.Copy(fixture.Output("B1number"), lost);
        using (var zip = ZipFile.Open(lost, ZipArchiveMode.Update)) zip.GetEntry("xl/calcChain.xml")!.Delete();
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, lost,
            [new CellExpectation("Sheet1", "B1", "number", "9")]),
            issue => issue.Gate == "G6" && issue.Code == "CALC_CHAIN_REMOVED");

        var retained = fixture.Output("chain-retained");
        File.Copy(fixture.Output("C1number"), retained);
        using (var zip = ZipFile.Open(retained, ZipArchiveMode.Update)) Add(zip, "xl/calcChain.xml", "stale");
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, retained,
            [new CellExpectation("Sheet1", "C1", "number", "9")]),
            issue => issue.Gate == "G6" && issue.Code == "STALE_CALC_CHAIN");
    }

    [Fact]
    public void VerifiedSaveReportsG6()
    {
        using var fixture = new Fixture();
        using var sessions = new ExcelSessions();
        var id = JsonSerializer.SerializeToElement(sessions.Open(fixture.Source)).GetProperty("session").GetString()!;
        sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "9")]);
        var saved = JsonSerializer.SerializeToElement(sessions.Save(id, fixture.Output("verified")));
        Assert.Contains(saved.GetProperty("gates").EnumerateArray(), gate => gate.GetString() == "G6");
        sessions.Close(id, true);
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void VerifiedSavePassesIndependentG6OnEverySyntheticFixture(string variant)
    {
        using var fixture = new Fixture();
        using var sessions = new ExcelSessions();
        var source = Path.Combine(fixture.Directory, variant + ".xlsx");
        var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
        sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "text", "new")]);
        var output = fixture.Output("checked");
        var saved = JsonSerializer.SerializeToElement(sessions.Save(id, output));
        Assert.Equal("verified", saved.GetProperty("status").GetString());
        Assert.Empty(AdvancedPartGate.Check(source, output,
            [new CellExpectation("Sheet1", "B1", "text", "new")]));
        sessions.Close(id, true);
    }

    [Fact]
    public void WorkbookContentTypeCannotBeChanged()
    {
        using var fixture = new Fixture();
        var output = fixture.Output("changed-type");
        File.Copy(fixture.Source, output);
        using (var zip = ZipFile.Open(output, ZipArchiveMode.Update))
            Modify(zip, "[Content_Types].xml", xml => xml.Replace("sheet.main+xml", "template.main+xml", StringComparison.Ordinal));
        Assert.Contains(AdvancedPartGate.Check(fixture.Source, output, []),
            issue => issue.Gate == "G6" && issue.Code == "CONTENT_TYPE_CHANGED");
    }

    [Fact]
    public void SignatureRelationshipWithoutSignaturePartStillBlocks()
    {
        using var fixture = new Fixture();
        using (var zip = ZipFile.Open(fixture.Source, ZipArchiveMode.Update))
            Modify(zip, "_rels/.rels", xml => xml.Replace("</Relationships>",
                "<Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"_xmlsignatures/origin.sigs\"/></Relationships>", StringComparison.Ordinal));
        Assert.Contains(AdvancedPartGate.CheckSignedSource(fixture.Source),
            issue => issue.Gate == "G6" && issue.Code == "SIGNED_PACKAGE_UNSUPPORTED");
    }

    [Fact]
    public void LocalBaselineDoesNotTriggerAdvancedPartFalsePositives()
    {
        var directory = Environment.GetEnvironmentVariable("DOCLOUPE_P2A_LOCAL_FIXTURES");
        if (string.IsNullOrWhiteSpace(directory)) return;
        var sources = Directory.GetFiles(directory, "*.*")
            .Where(path => Path.GetExtension(path) is ".xlsx" or ".xlsm")
            .Where(path => !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal))
            .OrderBy(path => path).ToArray();
        Assert.Equal(8, sources.Length);
        foreach (var source in sources)
        {
            var issues = AdvancedPartGate.Check(source, source, []);
            if (Path.GetFileName(source).StartsWith("05-", StringComparison.Ordinal))
                Assert.Contains(issues, issue => issue.Code == "SIGNED_PACKAGE_UNSUPPORTED");
            else Assert.Empty(issues);
        }
    }

    private static void AddCalculationChain(string path)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        Add(zip, "xl/calcChain.xml", "<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"C1\" i=\"1\"/></calcChain>");
        Modify(zip, "xl/_rels/workbook.xml.rels", xml => xml.Replace("</Relationships>",
            $"<Relationship Id=\"rId9\" Type=\"{Office}/calcChain\" Target=\"calcChain.xml\"/></Relationships>", StringComparison.Ordinal));
        Modify(zip, "[Content_Types].xml", xml => xml.Replace("</Types>",
            "<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/></Types>", StringComparison.Ordinal));
    }

    private static void Modify(ZipArchive zip, string part, Func<string, string> change)
    {
        var entry = zip.GetEntry(part)!;
        string text;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) text = reader.ReadToEnd();
        entry.Delete();
        Replace(zip, part, change(text));
    }

    private static void Replace(ZipArchive zip, string part, string text)
    {
        zip.GetEntry(part)?.Delete();
        Add(zip, part, text);
    }

    private static void Add(ZipArchive zip, string part, string text)
    {
        using var stream = zip.CreateEntry(part).Open();
        stream.Write(Encoding.UTF8.GetBytes(text));
    }

    private sealed class Fixture : IDisposable
    {
        public string Directory { get; } = Path.Combine(Path.GetTempPath(), "docloupe-g6-" + Guid.NewGuid().ToString("N"));
        public string Source => Path.Combine(Directory, "default.xlsx");
        public Fixture() => SyntheticFixtures.Create(Directory);
        public string Output(string name) => Path.Combine(Directory, name + ".xlsx");
        public void Dispose() => System.IO.Directory.Delete(Directory, true);
    }
}
