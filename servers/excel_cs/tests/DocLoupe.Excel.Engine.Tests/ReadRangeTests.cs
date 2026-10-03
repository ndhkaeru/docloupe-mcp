using System.IO.Compression;
using System.Text.Json;
using System.Xml;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ReadRangeTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void ReadsBoundedRectangleFromCurrentRevision(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-range-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var initial = Read(sessions, session, "Sheet1!A1:C2", "WrongDefault");
            Assert.Equal(0, initial.GetProperty("revision").GetInt32());
            Assert.Equal("hello", Value(initial, "A1"));
            Assert.Equal("42", Value(initial, "B1"));
            Assert.Equal("1+1", Formula(initial, "C1"));
            Assert.Equal(3, initial.GetProperty("cells").GetArrayLength());
            sessions.Apply(session, 0, [new SetValueOp("Sheet1", "B2", "number", "27")]);
            var preview = Read(sessions, session, "B1:C2", "Sheet1");
            var listed = JsonSerializer.SerializeToElement(sessions.Read(session, "Sheet1", ["A1", "B1:C2"]));
            Assert.Equal("hello", Value(listed, "A1"));
            Assert.Equal("27", Value(listed, "B2"));
            Assert.Equal(1, preview.GetProperty("revision").GetInt32());
            Assert.Equal("27", Value(preview, "B2"));
            Assert.DoesNotContain(preview.GetProperty("cells").EnumerateArray(),
                cell => cell.GetProperty("Address").GetString() == "A1");
            sessions.Close(session, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void MissingTargetUsesExplicitCellRangeAndCurrentRevision(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-used-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var cells = JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", []));
            Assert.Equal(4, cells.GetProperty("cells").GetArrayLength());
            Assert.Equal("old", Value(cells, "D3"));
            var values = JsonSerializer.SerializeToElement(sessions.Read(id, null, [], view: "values"));
            Assert.Equal(3, values.GetProperty("rows").GetArrayLength());
            Assert.Equal(4, values.GetProperty("rows")[0].GetArrayLength());
            Assert.Equal("old", values.GetProperty("rows")[2][3].GetString());
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "E4", "number", "77")]);
            var markdown = JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", [], view: "markdown"));
            Assert.Equal(1, markdown.GetProperty("revision").GetInt32());
            Assert.Contains("| 4 |", markdown.GetProperty("markdown").GetString());
            Assert.Contains("77", markdown.GetProperty("markdown").GetString());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void EmptyWorksheetHasEmptyDefaultReadInAllViews()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-used-empty-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
                var document = new XmlDocument();
                using (var stream = entry.Open()) document.Load(stream);
                var sheetData = document.DocumentElement!.GetElementsByTagName("sheetData",
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main")[0]!;
                sheetData.RemoveAll();
                entry.Delete();
                using var writer = new StreamWriter(archive.CreateEntry("xl/worksheets/sheet1.xml").Open());
                writer.Write(document.OuterXml);
            }
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, null, []))
                .GetProperty("cells").EnumerateArray());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, null, [], view: "values"))
                .GetProperty("rows").EnumerateArray());
            Assert.Equal("", JsonSerializer.SerializeToElement(sessions.Read(id, null, [], view: "markdown"))
                .GetProperty("markdown").GetString());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ReadCanIncludeMissingCellsInRequestedOrder()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-blanks-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            var cells = JsonSerializer.SerializeToElement(sessions.Read(session, "Sheet1", ["B2", "A1:B2"], false))
                .GetProperty("cells").EnumerateArray().ToArray();
            Assert.Equal(new[] { "B2", "A1", "B1", "A2", "B2" },
                cells.Select(cell => cell.GetProperty("Address").GetString()));
            Assert.All(cells.Where(cell => cell.GetProperty("Address").GetString() is "B2" or "A2"),
                cell =>
                {
                    Assert.Equal("blank", cell.GetProperty("Kind").GetString());
                    Assert.Equal(JsonValueKind.Null, cell.GetProperty("Value").ValueKind);
                });
            Assert.Equal("hello", cells[1].GetProperty("Value").GetString());
            Assert.Single(JsonSerializer.SerializeToElement(sessions.Read(session, "Sheet1", ["A1:B2"]))
                .GetProperty("cells").EnumerateArray(), cell => cell.GetProperty("Address").GetString() == "A1");
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("B2:A2")]
    [InlineData("B1:A2")]
    [InlineData("A1:B1:C1")]
    [InlineData("A1:Sheet1!B1")]
    [InlineData("A1:A501")]
    [InlineData("XFD1048576:XFD1048577")]
    public void RejectsInvalidOrTooLargeRanges(string target)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-reject-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.ThrowsAny<Exception>(() => sessions.Read(session, "Sheet1", [target]));
            Assert.Equal("hello", Value(Read(sessions, session, "A1:A500", "Sheet1"), "A1"));
            Assert.Equal(0, JsonSerializer.SerializeToElement(sessions.Status(session)).GetProperty("revision").GetInt32());
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void RejectsMixedSheetTargetsWithoutReadingAnyWorkbook()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-mixed-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var session = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.Throws<ArgumentException>(() => sessions.Read(session, "Sheet1", ["A1:B1", "Other!A1"]));
            sessions.Close(session, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static JsonElement Read(ExcelSessions sessions, string session, string target, string sheet) =>
        JsonSerializer.SerializeToElement(sessions.Read(session, sheet, [target]));

    private static string? Value(JsonElement result, string address) => result.GetProperty("cells").EnumerateArray()
        .Single(cell => cell.GetProperty("Address").GetString() == address).GetProperty("Value").GetString();

    private static string? Formula(JsonElement result, string address) => result.GetProperty("cells").EnumerateArray()
        .Single(cell => cell.GetProperty("Address").GetString() == address).GetProperty("Formula").GetString();
}
