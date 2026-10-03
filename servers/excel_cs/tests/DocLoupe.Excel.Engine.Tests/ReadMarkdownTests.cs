using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ReadMarkdownTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void MarkdownUsesAbsoluteHeadersAndFormulaAnnotations(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-read-markdown-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            var data = JsonSerializer.SerializeToElement(sessions.Read(id, null, ["Sheet1!B1:D3"], view: "markdown"));
            var markdown = data.GetProperty("markdown").GetString()!.Replace("\r\n", "\n", StringComparison.Ordinal);
            Assert.Equal("markdown", data.GetProperty("view").GetString());
            Assert.Contains("| row | B | C | D |", markdown);
            Assert.Contains("| 1 | 42 | 2 ƒ =1+1 |  |", markdown);
            Assert.Contains("| 2 |  |  |  |", markdown);
            Assert.Contains("| 3 |  |  | old |", markdown);
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "text", "a|b\\c<d>\nline")]);
            var changed = JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["B1"], view: "markdown"));
            Assert.Equal(1, changed.GetProperty("revision").GetInt32());
            Assert.Contains(@"a\|b\\c&lt;d&gt;<br>line", changed.GetProperty("markdown").GetString());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void MarkdownRejectsMultipleTargetsAndUnknownViews()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-markdown-shape-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            Assert.Throws<NotSupportedException>(() => sessions.Read(id, "Sheet1", ["A1", "B1"], view: "markdown"));
            Assert.Throws<NotSupportedException>(() => sessions.Read(id, "Sheet1", ["A1"], view: "full"));
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }
}
