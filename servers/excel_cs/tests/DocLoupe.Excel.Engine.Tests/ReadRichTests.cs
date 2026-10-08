using System.Text;
using System.Text.Json;
using System.Xml;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ReadRichTests
{
    [Theory]
    [InlineData("new-shared-strings")]
    [InlineData("prefixed-x")]
    public void RichProjectionReadsCurrentAndSavedRevisionsWithoutExposingXml(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-rich-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            const string markup = "<r>ệ &amp; &lt; </r><r b=\"false\" color=\"336699\"> text</r>";
            var operation = new SetValueOp("Sheet1", "E12", "inline", "ệ & <  text",
                Operation: "rich_set", RichMarkup: markup);
            sessions.Apply(session, 0, [operation, new SetValueOp("Sheet1", "F12", "text", "plain")]);
            var read = JsonSerializer.SerializeToElement(sessions.Read(session, "Sheet1", ["E12", "F12", "B2"], false,
                include: ["rich"]));
            Assert.Equal(1, read.GetProperty("revision").GetInt32());
            var cells = read.GetProperty("cells").EnumerateArray().ToArray();
            Assert.Equal("E12", cells[0].GetProperty("Address").GetString());
            Assert.Equal("<r>ệ &amp; &lt; </r><r b=\"false\" color=\"FF336699\"> text</r>",
                cells[0].GetProperty("rich").GetString());
            Assert.Equal("ệ & <  text", cells[0].GetProperty("Value").GetString());
            Assert.Equal("plain", cells[1].GetProperty("Value").GetString());
            Assert.Equal(JsonValueKind.Null, cells[1].GetProperty("rich").ValueKind);
            Assert.Equal(JsonValueKind.Null, cells[2].GetProperty("rich").ValueKind);
            Assert.DoesNotContain("CellMarkup", read.GetRawText());
            Assert.DoesNotContain("SharedMarkup", read.GetRawText());
            Assert.False(JsonSerializer.SerializeToElement(sessions.Read(session, "Sheet1", ["E12"]))
                .GetProperty("cells")[0].TryGetProperty("rich", out _));
            var output = Path.Combine(directory, "rich.xlsx");
            sessions.Save(session, output);
            var reopened = JsonSerializer.SerializeToElement(sessions.Open(output)).GetProperty("session").GetString()!;
            var saved = JsonSerializer.SerializeToElement(sessions.Read(reopened, "Sheet1", ["E12"],
                include: ["rich"]));
            Assert.Equal(cells[0].GetProperty("rich").GetString(),
                saved.GetProperty("cells")[0].GetProperty("rich").GetString());
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    public void ExistingSharedRunsAreProjectedWithExplicitFontProperties(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-shared-rich-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "source.xlsx");
            using (var store = new PackageStore(Path.Combine(directory, variant + ".xlsx")))
            {
                var shared = PackageStore.Parse(store.Read("xl/sharedStrings.xml"));
                foreach (var name in new[] { "rPh", "phoneticPr" })
                    foreach (var child in shared.GetElementsByTagName(name, PackageStore.Main).OfType<XmlElement>().ToArray())
                        child.ParentNode!.RemoveChild(child);
                store.Set("xl/sharedStrings.xml", Encoding.UTF8.GetBytes(shared.OuterXml));
                store.Save(source);
            }
            var rich = Path.Combine(directory, "shared-rich.xlsx");
            RichFixture.Create(source, rich, inline: false);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(rich)).GetProperty("session").GetString()!;
            var projected = JsonSerializer.SerializeToElement(sessions.Read(session, "Sheet1", ["A1"],
                include: ["rich"])).GetProperty("cells")[0].GetProperty("rich").GetString()!;
            Assert.Equal("<r>hello</r><r b=\"true\" color=\"FFFF0000\"> bold</r>", projected);
            Assert.True(RichTextAssertions.Matches(Assert.Single(P2aGates.ReadCells(rich, "Sheet1", ["A1"])), projected));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("<r><t>ok</t></r><!--hidden-->")]
    [InlineData("<r><t>ok</t><!--hidden--></r>")]
    [InlineData("<r><rPr><b/></rPr><t>ok</t></r><rPh sb=\"0\" eb=\"2\"><t>ok</t></rPh>")]
    [InlineData("<r><rPr><unknown/></rPr><t>ok</t></r>")]
    [InlineData("<r><t>ok<!--hidden--></t></r>")]
    public void UnknownSharedRichMarkupIsNotSilentlyFlattened(string content)
    {
        var cell = new CellRead("A1", "text", "ok", null,
            SharedMarkup: "<si xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                content + "</si>");
        Assert.Throws<NotSupportedException>(() => RichTextAssertions.ReadMarkup(cell));
    }

    [Fact]
    public void RichProjectionFailsClosedOnPhoneticAndInvalidInclude()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-rich-invalid-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.Throws<NotSupportedException>(() => sessions.Read(session, "Sheet1", ["A1"], include: ["rich"]));
            Assert.Throws<NotSupportedException>(() => sessions.Read(session, "Sheet1", ["A1"], include: ["style"]));
            Assert.Throws<NotSupportedException>(() => sessions.Read(session, "Sheet1", ["A1"], include: ["rich", "rich"]));
            Assert.Throws<NotSupportedException>(() => sessions.Read(session, "Sheet1", ["A1"],
                view: "values", include: ["rich"]));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(session))
                .GetProperty("revision").GetInt32());
        }
        finally { Directory.Delete(directory, true); }
    }
}
