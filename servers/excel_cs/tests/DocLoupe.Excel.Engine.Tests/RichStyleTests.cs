using System.Text;
using System.Text.Json;
using System.Xml;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class RichStyleTests
{
    private const string Original = "<r>hello</r><r b color=\"FFFF0000\"> bold</r>";
    private const string Italic = "<r i=\"true\">hello</r><r b i=\"true\" color=\"FFFF0000\"> bold</r>";

    [Fact]
    public void RichStyleAllUpdatesOnlySupportedRunsAndVerifiesSave()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-style-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "inline-rich.xlsx");
            RichFixture.Create(Path.Combine(directory, "new-shared-strings.xlsx"), source, inline: true);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            var operation = Request(JsonSerializer.Serialize(new
            {
                op = "rich_style", target = "Sheet1!A1", at = "all", style = new { italic = true },
                expect = new { rich = Original }
            }));
            sessions.Apply(id, 0, [operation], dryRun: true);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Apply(id, 0, [operation]);
            var output = Path.Combine(directory, "styled.xlsx");
            var assertion = new ValueAssertion("Sheet1", "A1", true, "text", "hello bold", null, Rich: Italic);
            Assert.Equal("verified", JsonSerializer.SerializeToElement(sessions.Save(id, output, [assertion]))
                .GetProperty("status").GetString());
            var cell = Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["A1"]));
            Assert.Equal("hello bold", cell.Value);
            Assert.True(RichTextAssertions.Matches(cell, Italic));
            Assert.Throws<NotSupportedException>(() => sessions.Apply(id, 1, [operation with { Expect = null }]));
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Undo(id, 1, 0);
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void RichStyleRejectsUnmodeledFormattingAndPhoneticsWithoutChangingRevision()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-style-guard-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "inline-rich.xlsx");
            RichFixture.Create(Path.Combine(directory, "new-shared-strings.xlsx"), source, inline: true);
            var unmodeled = Path.Combine(directory, "unmodeled.xlsx");
            using (var store = new PackageStore(source))
            {
                var part = store.SheetPart("Sheet1");
                var document = PackageStore.Parse(store.Read(part));
                var properties = Assert.Single(document.GetElementsByTagName("rPr", PackageStore.Main).OfType<XmlElement>());
                properties.AppendChild(document.CreateElement(properties.Prefix, "sz", PackageStore.Main));
                store.Set(part, Encoding.UTF8.GetBytes(document.OuterXml));
                store.Save(unmodeled);
            }
            var operation = Request("""{"op":"rich_style","target":"Sheet1!A1","at":"all","style":{"bold":false}}""");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(unmodeled)).GetProperty("session").GetString()!;
            Assert.Throws<NotSupportedException>(() => sessions.Apply(id, 0, [operation]));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, true);
            var phoneticId = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.Throws<NotSupportedException>(() => sessions.Apply(phoneticId, 0, [operation]));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(phoneticId)).GetProperty("revision").GetInt32());
            sessions.Close(phoneticId, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void RichStyleOnSharedRichReplaysAcrossUnrelatedRevisions()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-shared-style-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var intermediate = Path.Combine(directory, "rich-phonetic.xlsx");
            var source = Path.Combine(directory, "shared-rich.xlsx");
            RichFixture.Create(Path.Combine(directory, "default.xlsx"), intermediate, inline: false);
            using (var store = new PackageStore(intermediate))
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
            sessions.Apply(id, 0, [Request("""{"op":"set_value","target":"Sheet1!B1","value":27}""")]);
            sessions.Apply(id, 1, [Request(JsonSerializer.Serialize(new
            {
                op = "rich_style", target = "Sheet1!A1", at = "all", style = new { italic = true },
                expect = new { rich = Original }
            }))]);
            var output = Path.Combine(directory, "styled.xlsx");
            sessions.Save(id, output, [new ValueAssertion("Sheet1", "A1", false, null, null, null, Rich: Italic)]);
            var cell = Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["A1"]));
            Assert.Equal("inline", cell.Kind);
            Assert.True(RichTextAssertions.Matches(cell, Italic));
            using var saved = new PackageStore(output);
            Assert.Equal("0", PackageStore.Parse(saved.Read("xl/sharedStrings.xml")).DocumentElement!.GetAttribute("count"));
            sessions.Undo(id, 2, 1);
            var reverted = Path.Combine(directory, "reverted.xlsx");
            sessions.Save(id, reverted, [new ValueAssertion("Sheet1", "A1", false, null, null, null, Rich: Original)]);
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("outside-text")]
    [InlineData("outside-comment")]
    [InlineData("run-comment")]
    [InlineData("property-comment")]
    public void RichStyleRejectsUnmodeledNodes(string placement)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-rich-style-nodes-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "inline-rich.xlsx");
            RichFixture.Create(Path.Combine(directory, "new-shared-strings.xlsx"), source, inline: true);
            var invalid = Path.Combine(directory, "extra-node.xlsx");
            using (var store = new PackageStore(source))
            {
                var part = store.SheetPart("Sheet1");
                var document = PackageStore.Parse(store.Read(part));
                var cell = Assert.Single(document.GetElementsByTagName("c", PackageStore.Main)
                    .OfType<XmlElement>(), item => item.GetAttribute("r") == "A1");
                var container = Assert.Single(cell.GetElementsByTagName("is", PackageStore.Main).OfType<XmlElement>());
                var runs = container.ChildNodes.OfType<XmlElement>().ToArray();
                switch (placement)
                {
                    case "outside-text": container.InsertBefore(document.CreateTextNode("unexpected"), runs[1]); break;
                    case "outside-comment": container.InsertBefore(document.CreateComment("unexpected"), runs[1]); break;
                    case "run-comment": runs[0].AppendChild(document.CreateComment("unexpected")); break;
                    case "property-comment": Assert.Single(runs[1].GetElementsByTagName("rPr", PackageStore.Main)
                        .OfType<XmlElement>()).AppendChild(document.CreateComment("unexpected")); break;
                }
                store.Set(part, Encoding.UTF8.GetBytes(document.OuterXml));
                store.Save(invalid);
            }
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(invalid)).GetProperty("session").GetString()!;
            Assert.Throws<NotSupportedException>(() => sessions.Apply(id, 0,
                [Request("""{"op":"rich_style","target":"Sheet1!A1","at":"all","style":{"italic":true}}""")]));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("""{"op":"rich_style","target":"Sheet1!A1","at":{"match":"bold"},"style":{"italic":true}}""")]
    [InlineData("""{"op":"rich_style","target":"Sheet1!A1","at":"all","style":{}}""")]
    [InlineData("""{"op":"rich_style","target":"Sheet1!A1","at":"all","style":{"color":"theme:5"}}""")]
    [InlineData("""{"op":"rich_style","target":"Sheet1!A1","at":"all","style":{"bold":null}}""")]
    [InlineData("""{"op":"rich_style","target":"Sheet1!A1","at":"all","style":{"underline":true}}""")]
    [InlineData("""{"op":"rich_style","target":"Sheet1!A1","at":"all","style":{"bold":true,"bold":false}}""")]
    [InlineData("""{"op":"set_value","target":"Sheet1!A1","value":1,"style":{"bold":true}}""")]
    public void RichStyleRejectsUnsupportedInput(string json) => Assert.ThrowsAny<Exception>(() => Request(json));

    private static SetValueOp Request(string json) =>
        Assert.Single(JsonSerializer.Deserialize<SetValueRequest>(json)!.NormalizeMany(null));
}
