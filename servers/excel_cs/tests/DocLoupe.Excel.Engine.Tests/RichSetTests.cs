using System.Text;
using System.Text.Json;
using System.Xml;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class RichSetTests
{
    [Theory]
    [InlineData("new-shared-strings", "A1", "hello")]
    [InlineData("prefixed-x", "D3", "old")]
    public void RichSetWritesRunsAndIndependentSaveGateReadsThem(string variant, string address, string before)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-set-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            const string markup = "<r>new</r><r b=\"true\" color=\"FF0000\"> bold</r>";
            var operation = Request(JsonSerializer.Serialize(new
            {
                op = "rich_set", target = "Sheet1!" + address, rich = markup, expect = new { text = before }
            }));
            sessions.Apply(id, 0, [operation], dryRun: true);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Apply(id, 0, [operation]);
            var output = Path.Combine(directory, "rich.xlsx");
            var assertion = new ValueAssertion("Sheet1", address, true, "text", "new bold", null, Rich: markup);
            Assert.Equal("verified", JsonSerializer.SerializeToElement(sessions.Save(id, output, [assertion]))
                .GetProperty("status").GetString());
            var cell = Assert.Single(P2aGates.ReadCells(output, "Sheet1", [address]));
            Assert.Equal("inline", cell.Kind);
            Assert.Equal("new bold", cell.Value);
            Assert.True(RichTextAssertions.Matches(cell, markup));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void RichSetConvertsPlainSharedStringWithoutChangingOriginalItem()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-shared-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "plain-shared.xlsx");
            using (var store = new PackageStore(Path.Combine(directory, "default.xlsx")))
            {
                var document = PackageStore.Parse(store.Read("xl/sharedStrings.xml"));
                foreach (var name in new[] { "rPh", "phoneticPr" })
                    foreach (var child in document.GetElementsByTagName(name, PackageStore.Main).OfType<XmlElement>().ToArray())
                        child.ParentNode!.RemoveChild(child);
                store.Set("xl/sharedStrings.xml", Encoding.UTF8.GetBytes(document.OuterXml));
                store.Save(source);
            }
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [Request("""{"op":"rich_set","target":"Sheet1!A1","rich":"<r>hello</r><r i=\"true\"> world</r>","expect":{"value":"hello"}}""")]);
            var output = Path.Combine(directory, "written.xlsx");
            sessions.Save(id, output);
            var cell = Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["A1"]));
            Assert.Equal("inline", cell.Kind);
            Assert.Equal("hello world", cell.Value);
            using var written = new PackageStore(output);
            var shared = PackageStore.Parse(written.Read("xl/sharedStrings.xml"));
            Assert.Equal("0", shared.DocumentElement!.GetAttribute("count"));
            Assert.Equal("1", shared.DocumentElement.GetAttribute("uniqueCount"));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void StructuredRunsRoundTripAndVerifyOnSavedFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-runs-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var operation = Request("""{"op":"rich_set","target":"Sheet1!A1","runs":[{"text":"<&ệ"},{"text":" bold","font":{"bold":true,"italic":false,"color":"ff0000"}}]}""");
            Assert.Equal("<&ệ bold", operation.Value);
            var parsed = RichSetMarkup.Parse(operation.RichMarkup!);
            Assert.Equal(2, parsed.Runs.Count);
            Assert.Equal(new RichSetRun(" bold", true, false, "FFFF0000"), parsed.Runs[1]);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "new-shared-strings.xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [operation]);
            var output = Path.Combine(directory, "runs.xlsx");
            var assertion = new ValueAssertion("Sheet1", "A1", true, "text", "<&ệ bold", null,
                Rich: operation.RichMarkup);
            Assert.Equal("verified", JsonSerializer.SerializeToElement(sessions.Save(id, output, [assertion]))
                .GetProperty("status").GetString());
            Assert.True(RichTextAssertions.Matches(Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["A1"])),
                operation.RichMarkup!));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("""{"op":"rich_set","target":"Sheet1!A1","runs":[]}""")]
    [InlineData("""{"op":"rich_set","target":"Sheet1!A1","runs":[{"text":"x","font":{"bold":"yes"}}]}""")]
    [InlineData("""{"op":"rich_set","target":"Sheet1!A1","runs":[{"text":"x","font":{"color":"theme:5"}}]}""")]
    [InlineData("""{"op":"rich_set","target":"Sheet1!A1","runs":[{"text":"x","font":{"strike":true}}]}""")]
    [InlineData("""{"op":"rich_set","target":"Sheet1!A1","runs":[{"text":"x","unknown":1}]}""")]
    [InlineData("""{"op":"rich_set","target":"Sheet1!A1","runs":[{"text":"x"}],"rich":"<r>x</r>"}""")]
    [InlineData("""{"op":"rich_set","target":"Sheet1!A1","runs":[{"text":"x"}],"rich":null}""")]
    [InlineData("""{"op":"rich_set","target":"Sheet1!A1","rich":"<r>x</r>","runs":null}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!A1","value":1,"runs":[{"text":"x"}]}""")]
    public void StructuredRunsFailClosedOnUnsupportedInputs(string json)
    {
        Assert.ThrowsAny<Exception>(() => Request(json));
    }

    [Fact]
    public void RichSetRejectsPhoneticSourceAndInvalidMarkupBeforeChangingRevision()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-guard-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            Assert.ThrowsAny<Exception>(() => Request("""{"op":"rich_set","target":"Sheet1!D3","rich":"<!DOCTYPE rich><r>x</r>"}"""));
            Assert.ThrowsAny<Exception>(() => Request("""{"op":"rich_set","target":"Sheet1!D3","rich":"<r><b>x</b></r>"}"""));
            Assert.ThrowsAny<Exception>(() => Request("""{"op":"rich_set","target":"Sheet1!D3","rich":"<r>x</r>","runs":[]}"""));
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.Throws<NotSupportedException>(() => sessions.Apply(id, 0,
                [Request("""{"op":"rich_set","target":"Sheet1!A1","rich":"<r>x</r>"}""")]));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G4DetectsLostRunFormattingEvenWhenPlainTextIsUnchanged()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-intent-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            const string markup = "<r>hello</r><r b=\"true\"> world</r>";
            var operation = Request(JsonSerializer.Serialize(new
            {
                op = "rich_set", target = "Sheet1!A1", rich = markup
            }));
            var output = Path.Combine(directory, "correct.xlsx");
            using (var store = new PackageStore(Path.Combine(directory, "new-shared-strings.xlsx")))
            {
                SetValueEngine.Apply(store, [operation]);
                store.Save(output);
            }
            var expected = new CellExpectation("Sheet1", "A1", "inline", "hello world", RichMarkup: markup);
            Assert.Empty(P2aGates.CheckIntent(output, [expected]));
            var altered = Path.Combine(directory, "lost-format.xlsx");
            using (var store = new PackageStore(output))
            {
                var part = store.SheetPart("Sheet1");
                var text = Encoding.UTF8.GetString(store.Read(part));
                Assert.Contains("<b val=\"1\"", text);
                store.Set(part, Encoding.UTF8.GetBytes(text.Replace("<b val=\"1\"", "<b val=\"0\"", StringComparison.Ordinal)));
                store.Save(altered);
            }
            Assert.Equal("hello world", Assert.Single(P2aGates.ReadCells(altered, "Sheet1", ["A1"])).Value);
            Assert.Contains(P2aGates.CheckIntent(altered, [expected]), issue => issue.Code == "INTENT_RICH_MISMATCH");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void RichSetCanCheckPreviousRunsAndUndoToTheirRevision()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-undo-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "new-shared-strings.xlsx")))
                .GetProperty("session").GetString()!;
            const string first = "<r>first</r><r b=\"true\"> run</r>";
            const string second = "<r>second</r><r i=\"true\"> run</r>";
            sessions.Apply(id, 0, [Request(JsonSerializer.Serialize(new
                { op = "rich_set", target = "Sheet1!A1", rich = first }))]);
            sessions.Apply(id, 1, [Request(JsonSerializer.Serialize(new
                { op = "rich_set", target = "Sheet1!A1", rich = second, expect = new { rich = first } }))]);
            var output = Path.Combine(directory, "second.xlsx");
            sessions.Save(id, output, [new ValueAssertion("Sheet1", "A1", false, null, null, null, Rich: second)]);
            Assert.True(RichTextAssertions.Matches(Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["A1"])), second));
            sessions.Undo(id, 2, 1);
            var undone = Path.Combine(directory, "first.xlsx");
            sessions.Save(id, undone, [new ValueAssertion("Sheet1", "A1", false, null, null, null, Rich: first)]);
            Assert.True(RichTextAssertions.Matches(Assert.Single(P2aGates.ReadCells(undone, "Sheet1", ["A1"])), first));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static SetValueOp Request(string json) =>
        Assert.Single(JsonSerializer.Deserialize<SetValueRequest>(json)!.NormalizeMany(null));
}
