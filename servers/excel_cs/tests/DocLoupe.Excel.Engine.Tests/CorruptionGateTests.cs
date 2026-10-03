using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Schema;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class CorruptionGateTests
{
    [Fact]
    public void G1RejectsDanglingRelationshipReference()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace("<sheetData", "<sheetData r:id=\"rId999\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", StringComparison.Ordinal));
        Assert.Contains(P2aGates.CheckPackage(broken, ["xl/worksheets/sheet1.xml"]), issue => issue.Gate == "G1" && issue.Code == "DANGLING_REFERENCE");
    }

    [Fact]
    public void G2RejectsIncorrectCellChildOrder()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => Regex.Replace(xml, @"<f>1\+1</f>\s*<v>2</v>", "<v>2</v><f>1+1</f>"));
        Assert.Contains(DetachedValidator.Check(fixture.Source, broken, ["xl/worksheets/sheet1.xml"]).Issues, issue => issue.Code == "NEW_SCHEMA_ERROR");
    }

    [Theory]
    [InlineData("default", "")]
    [InlineData("prefixed-x", "x:")]
    public void G2ReportsMaskedParentWhenBaselineCellErrorExists(string variant, string prefix)
    {
        using var fixture = new Fixture();
        var source = Path.Combine(fixture.Directory, variant + ".xlsx");
        var invalid = fixture.Corrupt("xl/worksheets/sheet1.xml", xml =>
            Regex.Replace(xml, "<" + prefix + "c r=\"B1\" t=\"n\">\\s*<" + prefix + "v>42</" + prefix + "v>",
                "<" + prefix + "c r=\"B1\" t=\"n\"><" + prefix + "v>42</" + prefix + "v><" + prefix + "f>3</" + prefix + "f>"), source);
        var written = Path.Combine(fixture.Directory, "masked.xlsx");
        using (var store = new PackageStore(invalid))
        {
            SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "7")]);
            store.Save(written);
        }
        Assert.Contains(DetachedValidator.Check(invalid, written, ["xl/worksheets/sheet1.xml"]).Gaps,
            issue => issue.Code == "G2_MASKED_BY_BASELINE_ERROR");
    }

    [Fact]
    public void G2ReportsChangesBelowBaselineErrorWithCollidingLocalNames()
    {
        using var fixture = new Fixture();
        const string part = "xl/worksheets/sheet1.xml";
        var source = fixture.Corrupt(part, xml =>
            Regex.Replace(xml.Replace("<sheetData>",
                "<sheetData><other:row xmlns:other=\"urn:other\"/>", StringComparison.Ordinal),
                "(<c r=\"B1\" t=\"n\">\\s*<v>42</v>)", "$1<f>3</f>"));
        var written = fixture.Corrupt(part, xml => xml.Replace("<v>42</v><f>3</f>",
            "<v>9</v><f>3</f>", StringComparison.Ordinal), source);
        Assert.Contains(DetachedValidator.CheckPackage(source).Issues,
            issue => issue.Part == part && issue.Detail.Contains("/x:c[2]", StringComparison.Ordinal));
        var result = DetachedValidator.Check(source, written, [part]);
        Assert.Empty(result.Issues);
        Assert.Contains(result.Gaps, issue => issue.Code == "G2_MASKED_BY_BASELINE_ERROR");
    }

    [Fact]
    public void G2DoesNotMaskUnrelatedCellWhenBaselineErrorIsLocal()
    {
        using var fixture = new Fixture();
        const string part = "xl/worksheets/sheet1.xml";
        var source = fixture.Corrupt(part, xml =>
            Regex.Replace(xml, "(<c r=\"B1\" t=\"n\">\\s*<v>42</v>)", "$1<f>3</f>"));
        var written = fixture.Corrupt(part, xml => xml.Replace("<f>1+1</f>",
            "<f>1+2</f>", StringComparison.Ordinal), source);
        var result = DetachedValidator.Check(source, written, [part]);
        Assert.Empty(result.Issues);
        Assert.Empty(result.Gaps);
    }

    [Fact]
    public void G3RejectsUndeclaredMcPrefix()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace("mc:Ignorable=\"x14ac\"", "mc:Ignorable=\"x14ac absent\"", StringComparison.Ordinal));
        Assert.Contains(P2aMarkupGate.Check(fixture.Source, broken, ["xl/worksheets/sheet1.xml"]), issue => issue.Gate == "G3" && issue.Code == "UNDECLARED_MC_PREFIX");
    }

    [Fact]
    public void G3RejectsPrefixRewriteWithinAnEditedCell()
    {
        using var fixture = new Fixture();
        var source = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace(
            "<worksheet xmlns=", "<worksheet xmlns:a=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns=",
            StringComparison.Ordinal));
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace(
            "<v>42</v>", "<a:v>42</a:v>", StringComparison.Ordinal), source);
        Assert.Empty(P2aGates.CheckTouchedCells(source, broken,
            [new CellExpectation("Sheet1", "B1", "number", "42")]));
        Assert.Contains(P2aMarkupGate.Check(source, broken, ["xl/worksheets/sheet1.xml"]),
            issue => issue.Gate == "G3" && issue.Code == "PREFIX_REWRITTEN");
    }

    [Fact]
    public void G4RejectsSilentLostEdit()
    {
        using var fixture = new Fixture();
        Assert.Contains(P2aGates.CheckIntent(fixture.Source, [new CellExpectation("Sheet1", "B1", "number", "999")]),
            issue => issue.Gate == "G4" && issue.Code == "INTENT_MISMATCH");
    }

    [Fact]
    public void G4ReportsMissingTargetInsteadOfTreatingItAsBlank()
    {
        using var fixture = new Fixture();
        Assert.Contains(P2aGates.CheckIntent(fixture.Source, [new CellExpectation("Sheet1", "B2", "number", "9")]),
            issue => issue.Gate == "G4" && issue.Code == "INTENT_MISSING");
    }

    [Fact]
    public void G5RejectsUnrequestedStyleInsideEditedCell()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace("<c r=\"B1\" t=\"n\">", "<c r=\"B1\" t=\"n\" s=\"5\">", StringComparison.Ordinal));
        Assert.Contains(P2aGates.CheckTouchedCells(fixture.Source, broken, [new CellExpectation("Sheet1", "B1", "number", "42")]),
            issue => issue.Gate == "G5" && issue.Code == "CELL_ATTRIBUTE_CHANGED");
    }

    [Fact]
    public void G5RejectsUnrequestedLocalNamespaceOnEditedCell()
    {
        using var fixture = new Fixture();
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace(
            "<c r=\"B1\" t=\"n\">", "<c r=\"B1\" t=\"n\" xmlns:foreign=\"urn:test\">", StringComparison.Ordinal));
        Assert.Empty(P2aGates.CheckIntent(broken, [new CellExpectation("Sheet1", "B1", "number", "42")]));
        Assert.Contains(P2aGates.CheckTouchedCells(fixture.Source, broken,
            [new CellExpectation("Sheet1", "B1", "number", "42")]),
            issue => issue.Gate == "G5" && issue.Code == "CELL_ATTRIBUTE_CHANGED");
    }

    [Theory]
    [InlineData("style")]
    [InlineData("child")]
    [InlineData("namespace")]
    public void G5RejectsUnmodeledContentInNewCell(string corruption)
    {
        using var fixture = new Fixture();
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        using (var store = new PackageStore(fixture.Source))
        {
            SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B2", "number", "100")]);
            store.Save(edited);
        }
        var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => corruption switch
        {
            "style" => xml.Replace("<c r=\"B2\"", "<c r=\"B2\" s=\"5\"", StringComparison.Ordinal),
            "namespace" => xml.Replace("<c r=\"B2\"", "<c r=\"B2\" xmlns:foreign=\"urn:test\"", StringComparison.Ordinal),
            _ => Regex.Replace(xml, "(<c r=\"B2\"[^>]*>)", "$1<foreign xmlns=\"urn:test\"/>")
        }, edited);
        Assert.Empty(P2aGates.CheckIntent(broken, [new CellExpectation("Sheet1", "B2", "number", "100")]));
        Assert.Contains(P2aGates.CheckTouchedCells(fixture.Source, broken,
            [new CellExpectation("Sheet1", "B2", "number", "100")]),
            issue => issue.Gate == "G5" && issue.Code is "CELL_ATTRIBUTE_CHANGED" or "CELL_UNMODELED_CHILD");
    }

    [Fact]
    public void G5RejectsUndeclaredBytesOutsideDeclaredEdit()
    {
        using var fixture = new Fixture();
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        using (var store = new PackageStore(fixture.Source))
        {
            var result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(edited);
            var broken = fixture.Corrupt("xl/worksheets/sheet1.xml", xml => xml.Replace("old", "damaged", StringComparison.Ordinal), edited);
            Assert.Contains(P2aGates.CheckPreservation(fixture.Source, broken,
                result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After)), []),
                issue => issue.Gate == "G5" && issue.Code.StartsWith("UNDECLARED", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void G5RejectsUnrelatedFormulaHiddenInsideWriterDeclaredSpan()
    {
        using var fixture = new Fixture();
        const string part = "xl/worksheets/sheet1.xml";
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        ApplyResult result;
        using (var store = new PackageStore(fixture.Source))
        {
            result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(edited);
        }
        var broken = fixture.Corrupt(part, xml => xml.Replace("<f>1+1</f>", "<f>9*9</f>", StringComparison.Ordinal), edited);
        static byte[] PartBytes(string path, string name)
        {
            using var archive = ZipFile.OpenRead(path);
            using var stream = archive.GetEntry(name)!.Open();
            using var output = new MemoryStream();
            stream.CopyTo(output);
            return output.ToArray();
        }
        static (int Start, int End) CellSpan(string xml)
        {
            var start = xml.IndexOf("<c r=\"B1\"", StringComparison.Ordinal);
            var formula = xml.IndexOf("<c r=\"C1\"", start, StringComparison.Ordinal);
            var end = xml.IndexOf("</c>", formula, StringComparison.Ordinal) + "</c>".Length;
            Assert.True(start >= 0 && formula > start && end > formula);
            return (Encoding.UTF8.GetByteCount(xml[..start]), Encoding.UTF8.GetByteCount(xml[..end]));
        }
        var before = PartBytes(fixture.Source, part);
        var after = PartBytes(broken, part);
        var (start, end) = CellSpan(Encoding.UTF8.GetString(before));
        var (writtenStart, writtenEnd) = CellSpan(Encoding.UTF8.GetString(after));
        var declared = result.Edits.Where(edit => !edit.Part.Equals(part, StringComparison.OrdinalIgnoreCase))
            .Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After))
            .Append(new DeclaredByteSpan(part, start, end, before[start..end], after[writtenStart..writtenEnd])).ToArray();
        var intent = new CellExpectation("Sheet1", "B1", "number", "100");
        var issues = new List<GateIssue>();
        issues.AddRange(P2aGates.CheckPackage(broken, result.ChangedParts));
        var schema = DetachedValidator.Check(fixture.Source, broken, result.ChangedParts);
        Assert.Empty(schema.Gaps);
        issues.AddRange(schema.Issues.Select(issue => new GateIssue("G2", issue.Code, issue.Detail)));
        issues.AddRange(P2aMarkupGate.Check(fixture.Source, broken, result.ChangedParts));
        issues.AddRange(P2aGates.CheckIntent(broken, [intent]));
        issues.AddRange(P2aGates.CheckPreservation(fixture.Source, broken, declared, []));
        issues.AddRange(P2aGates.CheckTouchedCells(fixture.Source, broken, [intent]));
        issues.AddRange(P2aGates.CheckSemanticPreservation(fixture.Source, broken, [intent], declared));
        Assert.Contains(issues, issue => issue.Gate == "G5" && issue.Code == "DECLARATION_OUTSIDE_INTENT");
        issues.AddRange(AdvancedPartGate.Check(fixture.Source, broken, [intent]));
        Assert.Contains(issues, issue => issue.Gate == "G5" && issue.Detail.Contains("C1", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("xl/workbook.xml", "calcId=\"191029\"", "calcId=\"1\"", "UNDECLARED_WORKBOOK_CHANGE")]
    [InlineData("xl/sharedStrings.xml", "count=\"1\"", "count=\"9\"", "UNDECLARED_SHARED_STRING_CHANGE")]
    [InlineData("xl/sharedStrings.xml", "<t>hello</t>", "<t>corrupted</t>", "UNDECLARED_SHARED_STRING_CHANGE")]
    [InlineData("xl/_rels/workbook.xml.rels", "/sharedStrings\"", "/styles\"", "UNDECLARED_RELATIONSHIP_CHANGE")]
    [InlineData("[Content_Types].xml", "spreadsheetml.sharedStrings+xml", "spreadsheetml.styles+xml", "UNDECLARED_CONTENT_TYPE_CHANGE")]
    [InlineData("_rels/.rels", "Id=\"rId1\"", "Id=\"rId9\"", "UNDECLARED_PART_CHANGE")]
    public void G5RejectsUnintendedChangesToOtherParts(string part, string from, string to, string code)
    {
        using var fixture = new Fixture();
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        using (var store = new PackageStore(fixture.Source))
        {
            SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(edited);
        }
        var broken = fixture.Corrupt(part, xml => xml.Replace(from, to, StringComparison.Ordinal), edited);
        Assert.Contains(P2aGates.CheckSemanticPreservation(fixture.Source, broken,
            [new CellExpectation("Sheet1", "B1", "number", "100")]), issue => issue.Code == code);
    }

    [Fact]
    public void G5RejectsFabricatedSharedStringsMetadataAndRelationship()
    {
        using var fixture = new Fixture();
        var source = Path.Combine(fixture.Directory, "new-shared-strings.xlsx");
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        var intent = new CellExpectation("Sheet1", "B1", "text", "new value");
        ApplyResult result;
        using (var store = new PackageStore(source))
        {
            result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "text", "new value")]);
            store.Save(edited);
        }
        var spans = result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End,
            edit.Before, edit.After)).ToArray();
        Assert.Empty(P2aGates.CheckSemanticPreservation(source, edited, [intent], spans));
        var count = fixture.Corrupt("xl/sharedStrings.xml", xml => xml.Replace("count=\"1\"", "count=\"9\"", StringComparison.Ordinal), edited);
        Assert.Contains(P2aGates.CheckSemanticPreservation(source, count, [intent]), issue => issue.Code == "UNDECLARED_SHARED_STRING_CHANGE");
        var type = fixture.Corrupt("[Content_Types].xml", xml => xml.Replace("spreadsheetml.sharedStrings+xml", "spreadsheetml.styles+xml", StringComparison.Ordinal), edited);
        Assert.Contains(P2aGates.CheckSemanticPreservation(source, type, [intent]), issue => issue.Code == "UNDECLARED_CONTENT_TYPE_CHANGE");
    }

    [Theory]
    [InlineData("xl/workbook.xml", "<calcPr ", "calcId=\"191029\"", "calcId='191029'")]
    [InlineData("xl/sharedStrings.xml", "<sst ", "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'")]
    public void G5RejectsUnexpectedLexicalChangesInsideAnOtherwisePermittedStartTag(
        string part, string tag, string from, string to)
    {
        using var fixture = new Fixture();
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        ApplyResult result;
        using (var store = new PackageStore(fixture.Source))
        {
            result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(edited);
        }
        var broken = fixture.Corrupt(part, xml => xml.Replace(from, to, StringComparison.Ordinal), edited);
        static byte[] Bytes(string path, string name)
        {
            using var archive = ZipFile.OpenRead(path);
            using var input = archive.GetEntry(name)!.Open();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }
        static (int Start, int End) TagSpan(byte[] bytes, string tag)
        {
            var xml = Encoding.UTF8.GetString(bytes);
            var start = xml.IndexOf(tag, StringComparison.Ordinal);
            Assert.True(start >= 0);
            var end = xml.IndexOf('>', start) + 1;
            return (Encoding.UTF8.GetByteCount(xml[..start]), Encoding.UTF8.GetByteCount(xml[..end]));
        }
        var original = Bytes(fixture.Source, part);
        var modified = Bytes(broken, part);
        var before = TagSpan(original, tag);
        var after = TagSpan(modified, tag);
        var spans = result.Edits.Where(edit => !edit.Part.Equals(part, StringComparison.OrdinalIgnoreCase))
            .Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After))
            .Append(new DeclaredByteSpan(part, before.Start, before.End,
                original[before.Start..before.End], modified[after.Start..after.End])).ToArray();
        Assert.Empty(P2aGates.CheckPreservation(fixture.Source, broken, spans, []));
        Assert.Contains(P2aGates.CheckSemanticPreservation(fixture.Source, broken,
            [new CellExpectation("Sheet1", "B1", "number", "100")], spans),
            issue => issue.Code == "DECLARATION_OUTSIDE_INTENT");
    }

    [Fact]
    public void G5RejectsUndeclaredWhitespaceInsertedBesideAnUnrelatedCell()
    {
        using var fixture = new Fixture();
        const string part = "xl/worksheets/sheet1.xml";
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        ApplyResult result;
        using (var store = new PackageStore(fixture.Source))
        {
            result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(edited);
        }
        var broken = fixture.Corrupt(part, xml => xml.Replace("<c r=\"C1\"", "  <c r=\"C1\"", StringComparison.Ordinal), edited);
        using var archive = ZipFile.OpenRead(fixture.Source);
        using var stream = archive.GetEntry(part)!.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var before = memory.ToArray();
        var offset = Encoding.UTF8.GetByteCount(Encoding.UTF8.GetString(before).Split("<c r=\"C1\"", 2)[0]);
        var spans = result.Edits.Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End,
            edit.Before, edit.After)).Append(new DeclaredByteSpan(part, offset, offset, [], Encoding.UTF8.GetBytes("  "))).ToArray();
        Assert.Empty(P2aGates.CheckPreservation(fixture.Source, broken, spans, []));
        Assert.Contains(P2aGates.CheckSemanticPreservation(fixture.Source, broken,
            [new CellExpectation("Sheet1", "B1", "number", "100")], spans),
            issue => issue.Code == "DECLARATION_OUTSIDE_INTENT");
    }

    [Theory]
    [InlineData("xl/sharedStrings.xml", "<t>hello</t>", "<t>&#104;ello</t>")]
    [InlineData("xl/workbook.xml", "calcId=\"191029\"", "calcId='191029'")]
    [InlineData("xl/_rels/workbook.xml.rels", "Id=\"rId1\"", "Id='rId1'")]
    [InlineData("[Content_Types].xml", "Extension=\"rels\"", "Extension='rels'")]
    public void G5RejectsBroadSpansInMetadataPartsEvenWhenXmlIsSemanticallyIdentical(
        string part, string from, string to)
    {
        using var fixture = new Fixture();
        var edited = Path.Combine(fixture.Directory, "edited.xlsx");
        ApplyResult result;
        using (var store = new PackageStore(fixture.Source))
        {
            result = SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(edited);
        }
        var broken = fixture.Corrupt(part, xml => xml.Replace(from, to, StringComparison.Ordinal), edited);
        static byte[] Bytes(string path, string name)
        {
            using var archive = ZipFile.OpenRead(path);
            using var input = archive.GetEntry(name)!.Open();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }
        var original = Bytes(fixture.Source, part);
        var modified = Bytes(broken, part);
        var spans = result.Edits.Where(edit => !edit.Part.Equals(part, StringComparison.OrdinalIgnoreCase))
            .Select(edit => new DeclaredByteSpan(edit.Part, edit.Start, edit.End, edit.Before, edit.After))
            .Append(new DeclaredByteSpan(part, 0, original.Length, original, modified)).ToArray();
        Assert.Empty(P2aGates.CheckPreservation(fixture.Source, broken, spans, []));
        Assert.Contains(P2aGates.CheckSemanticPreservation(fixture.Source, broken,
            [new CellExpectation("Sheet1", "B1", "number", "100")], spans),
            issue => issue.Gate == "G5" && issue.Code == "DECLARATION_OUTSIDE_INTENT");
    }

    [Fact]
    public void WriterAcceptsUtf8DeclarationWithoutEncoding()
    {
        using var fixture = new Fixture();
        var source = fixture.Corrupt("xl/worksheets/sheet1.xml", xml =>
            xml.Replace(" encoding=\"UTF-8\"", "", StringComparison.Ordinal));
        var output = Path.Combine(fixture.Directory, "without-encoding.xlsx");
        using (var store = new PackageStore(source))
        {
            SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(output);
        }
        Assert.Empty(P2aGates.CheckIntent(output, [new CellExpectation("Sheet1", "B1", "number", "100")]));
    }

    [Fact]
    public void WorkbookCalculationFlagIsNotInsertedIntoAnAttributeValue()
    {
        using var fixture = new Fixture();
        var source = fixture.Corrupt("xl/workbook.xml", xml =>
            xml.Replace("<calcPr calcId=", "<calcPr xmlns:p=\"urn:calcPr\" calcId=", StringComparison.Ordinal));
        var output = Path.Combine(fixture.Directory, "calculation.xlsx");
        using (var store = new PackageStore(source))
        {
            SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "100")]);
            store.Save(output);
        }
        using var archive = ZipFile.OpenRead(output);
        using var stream = archive.GetEntry("xl/workbook.xml")!.Open();
        var document = new System.Xml.XmlDocument();
        document.Load(stream);
        var calculation = document.GetElementsByTagName("calcPr", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")[0]!;
        Assert.Equal("1", calculation.Attributes!["fullCalcOnLoad"]!.Value);
        Assert.Equal("urn:calcPr", calculation.Attributes["xmlns:p"]!.Value);
    }

    private sealed class Fixture : IDisposable
    {
        public string Directory { get; } = Path.Combine(Path.GetTempPath(), "docloupe-corrupt-" + Guid.NewGuid().ToString("N"));
        public string Source { get; }
        public Fixture()
        {
            SyntheticFixtures.Create(Directory);
            Source = Path.Combine(Directory, "default.xlsx");
        }
        public string Corrupt(string part, Func<string, string> transform, string? original = null)
        {
            original ??= Source;
            var result = Path.Combine(Directory, Guid.NewGuid().ToString("N") + ".xlsx");
            using var archive = ZipFile.OpenRead(original);
            using var destination = new ZipArchive(File.Create(result), ZipArchiveMode.Create);
            foreach (var entry in archive.Entries)
            {
                var output = destination.CreateEntry(entry.FullName);
                using var input = entry.Open();
                using var stream = output.Open();
                if (entry.FullName == part)
                {
                    using var reader = new StreamReader(input, Encoding.UTF8);
                    var old = reader.ReadToEnd();
                    var modified = transform(old);
                    Assert.NotEqual(old, modified);
                    stream.Write(Encoding.UTF8.GetBytes(modified));
                }
                else input.CopyTo(stream);
            }
            return result;
        }
        public void Dispose() => System.IO.Directory.Delete(Directory, true);
    }
}
