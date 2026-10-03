using System.IO.Compression;
using System.Text.Json;
using System.Xml;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class FindTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void BoundedSearchUsesCurrentSessionAndNfcNormalization(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-find-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var initial = JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "A1:D3", "HELLO", false));
            Assert.True(initial.GetProperty("partial").GetBoolean());
            Assert.Equal(12, initial.GetProperty("total_scanned").GetInt32());
            Assert.Equal("Sheet1!A1", Assert.Single(initial.GetProperty("matches").EnumerateArray()).GetProperty("addr").GetString());
            Assert.False(initial.GetProperty("truncated").GetBoolean());
            var wholeWorkbook = JsonSerializer.SerializeToElement(sessions.Find(id, null, null, "HELLO", false));
            Assert.Equal("Sheet1!A1", Assert.Single(wholeWorkbook.GetProperty("matches").EnumerateArray())
                .GetProperty("addr").GetString());
            Assert.Equal(12, wholeWorkbook.GetProperty("total_scanned").GetInt32());
            var sheetOnly = JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", null, "hello", false));
            Assert.Single(sheetOnly.GetProperty("matches").EnumerateArray());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "A1:D3", "HELLO", false,
                caseSensitive: true)).GetProperty("matches").EnumerateArray());
            Assert.Equal("Sheet1!C1", Assert.Single(JsonSerializer.SerializeToElement(sessions.Find(id, null,
                "Sheet1!A1:D3", "1\\+1", true, searchIn: "formula")).GetProperty("matches").EnumerateArray())
                .GetProperty("addr").GetString());
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "E4", "text", "cafe\u0301")]);
            var normalized = JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "E4", "café", false));
            Assert.Equal(1, normalized.GetProperty("revision").GetInt32());
            Assert.Single(normalized.GetProperty("matches").EnumerateArray());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "E4", "café", false,
                normalize: "none")).GetProperty("matches").EnumerateArray());
            Assert.False(normalized.GetProperty("truncated").GetBoolean());
            sessions.Apply(id, 1, [new SetValueOp("Sheet1", "F4", "text", "cafe\u0301")]);
            var limited = JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "E4:F4", "café", false,
                maxResults: 1));
            Assert.Single(limited.GetProperty("matches").EnumerateArray());
            Assert.True(limited.GetProperty("truncated").GetBoolean());
            Assert.Equal(2, limited.GetProperty("total_scanned").GetInt32());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void UnscopedSearchTraversesAllSheetsInWorkbookOrder()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-find-sheets-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var path = Path.Combine(directory, "default.xlsx");
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                using (var original = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open())
                using (var clone = archive.CreateEntry("xl/worksheets/sheet2.xml").Open())
                    original.CopyTo(clone);
                EditXml(archive, "xl/workbook.xml", document =>
                {
                    var sheet = (XmlElement)document.GetElementsByTagName("sheet",
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main")[0]!;
                    var copy = (XmlElement)sheet.CloneNode(true);
                    copy.SetAttribute("name", "Sheet2");
                    copy.SetAttribute("sheetId", "2");
                    copy.SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId3");
                    sheet.ParentNode!.AppendChild(copy);
                });
                EditXml(archive, "xl/_rels/workbook.xml.rels", document =>
                {
                    var relationship = document.CreateElement("Relationship",
                        "http://schemas.openxmlformats.org/package/2006/relationships");
                    relationship.SetAttribute("Id", "rId3");
                    relationship.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                    relationship.SetAttribute("Target", "worksheets/sheet2.xml");
                    document.DocumentElement!.AppendChild(relationship);
                });
                EditXml(archive, "[Content_Types].xml", document =>
                {
                    var part = document.CreateElement("Override",
                        "http://schemas.openxmlformats.org/package/2006/content-types");
                    part.SetAttribute("PartName", "/xl/worksheets/sheet2.xml");
                    part.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                    document.DocumentElement!.AppendChild(part);
                });
            }
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(path)).GetProperty("session").GetString()!;
            var found = JsonSerializer.SerializeToElement(sessions.Find(id, null, null, "hello", false));
            Assert.Equal(new[] { "Sheet1!A1", "Sheet2!A1" }, found.GetProperty("matches").EnumerateArray()
                .Select(item => item.GetProperty("addr").GetString()).ToArray());
            Assert.Equal(24, found.GetProperty("total_scanned").GetInt32());
            Assert.Single(JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet2", null, "hello", false))
                .GetProperty("matches").EnumerateArray());
            var limited = JsonSerializer.SerializeToElement(sessions.Find(id, null, null, "hello", false,
                maxResults: 1));
            Assert.True(limited.GetProperty("truncated").GetBoolean());
            Assert.Equal(13, limited.GetProperty("total_scanned").GetInt32());
            sessions.Apply(id, 0, [new SetValueOp("Sheet2", "A1000", "number", "1")]);
            Assert.Throws<ArgumentException>(() => sessions.Find(id, null, null, "hello", false));
            Assert.Single(JsonSerializer.SerializeToElement(sessions.Find(id, "Sheet1", "A1:D3", "hello", false))
                .GetProperty("matches").EnumerateArray());
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static void EditXml(ZipArchive archive, string path, Action<XmlDocument> edit)
    {
        var entry = archive.GetEntry(path)!;
        var document = new XmlDocument();
        using (var stream = entry.Open()) document.Load(stream);
        edit(document);
        entry.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(path).Open());
        writer.Write(document.OuterXml);
    }

    [Fact]
    public void SearchLimitsAndQueryShapeFailClosed()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-find-reject-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            foreach (var json in new[]
                {
                    "{}", "{\"text\":\"x\",\"regex\":\"y\"}",
                    "{\"text\":\"x\",\"style\":{}}", "{\"regex\":\"\"}"
                })
                Assert.ThrowsAny<Exception>(() => JsonSerializer.Deserialize<FindQueryRequest>(json)!.Normalize());
            Assert.ThrowsAny<Exception>(() => sessions.Find(id, "Sheet1", "A1:A501", "x", false));
            Assert.Throws<KeyNotFoundException>(() => sessions.Find(id, "NotASheet", null, "x", false));
            Assert.Throws<NotSupportedException>(() => sessions.Find(id, "Sheet1", "A1", "x", false, searchIn: "display"));
            Assert.ThrowsAny<Exception>(() => sessions.Find(id, "Sheet1", "A1", "[", true));
            Assert.ThrowsAny<Exception>(() => sessions.Find(id, "Sheet1", "A1", "x", false, maxResults: 101));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("revision").GetInt32());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }
}
