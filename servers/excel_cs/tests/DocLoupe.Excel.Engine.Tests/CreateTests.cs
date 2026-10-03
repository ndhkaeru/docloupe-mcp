using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class CreateTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void TemplateCreationCopiesBytesAndOpensIndependentSession(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-create-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var template = Path.Combine(directory, variant + ".xlsx");
            var original = File.ReadAllBytes(template);
            var created = Path.Combine(directory, "created.xlsx");
            using var sessions = new ExcelSessions();
            var response = JsonSerializer.SerializeToElement(sessions.CreateFromTemplate(template, created));
            var id = response.GetProperty("session").GetString()!;
            Assert.True(response.GetProperty("new").GetBoolean());
            Assert.Equal(created, response.GetProperty("default_path").GetString());
            Assert.Equal(original, File.ReadAllBytes(created));
            Assert.Equal("42", Assert.Single(P2aGates.ReadCells(created, "Sheet1", ["B1"])).Value);
            Assert.False(JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("dirty").GetBoolean());
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            var output = Path.Combine(directory, "saved.xlsx");
            sessions.Save(id, output);
            Assert.Equal("27", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            Assert.Equal(original, File.ReadAllBytes(template));
            Assert.Equal(original, File.ReadAllBytes(created));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsm")]
    [InlineData("xltx")]
    [InlineData("xltm")]
    public void NewWorkbookCreatesIndependentSheetsAndPassesTheFirstSave(string format)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-create-new-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var created = Path.Combine(directory, "fresh." + format);
            using var sessions = new ExcelSessions();
            var response = JsonSerializer.SerializeToElement(sessions.CreateNew(created,
                ["Sheet1", "Dữ liệu"], "Dữ liệu"));
            var id = response.GetProperty("session").GetString()!;
            Assert.Equal(["Sheet1", "Dữ liệu"], response.GetProperty("sheets").EnumerateArray()
                .Select(item => item.GetString()!).ToArray());
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Read(id, "Dữ liệu", []))
                .GetProperty("cells").EnumerateArray());
            Assert.False(JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("dirty").GetBoolean());
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B2", "number", "3"),
                new SetValueOp("Dữ liệu", "A1", "text", "mới")]);
            var output = Path.Combine(directory, "saved." + format);
            var report = JsonSerializer.SerializeToElement(sessions.Save(id, output));
            Assert.Equal("verified", report.GetProperty("status").GetString());
            Assert.Equal("3", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["B2"])).Value);
            Assert.Equal("mới", Assert.Single(P2aGates.ReadCells(output, "Dữ liệu", ["A1"])).Value);
            Assert.Empty(P2aGates.ReadCells(created, "Dữ liệu", ["A1"]));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsm")]
    [InlineData("xltx")]
    [InlineData("xltm")]
    public void TemplateCreationPreservesEveryWorkbookFormat(string format)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-create-format-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var original = Path.Combine(directory, "source." + format);
            var copied = Path.Combine(directory, "copied." + format);
            using var sessions = new ExcelSessions();
            var initial = JsonSerializer.SerializeToElement(sessions.CreateNew(original));
            sessions.Close(initial.GetProperty("session").GetString()!, false);
            var created = JsonSerializer.SerializeToElement(sessions.CreateFromTemplate(original, copied));
            Assert.Equal(File.ReadAllBytes(original), File.ReadAllBytes(copied));
            Assert.Equal(copied, created.GetProperty("path").GetString());
            sessions.Close(created.GetProperty("session").GetString()!, false);
            Assert.Throws<NotSupportedException>(() => sessions.CreateFromTemplate(original,
                Path.Combine(directory, format == "xlsx" ? "different.xlsm" : "different.xlsx")));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void NewWorkbookRejectsInvalidSheetNamesAndExistingOutput()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-create-new-errors-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var output = Path.Combine(directory, "fresh.xlsx");
            using var sessions = new ExcelSessions();
            Assert.ThrowsAny<Exception>(() => sessions.CreateNew(output, []));
            Assert.ThrowsAny<Exception>(() => sessions.CreateNew(output, ["ABC", "abc"]));
            Assert.ThrowsAny<Exception>(() => sessions.CreateNew(output, ["Invalid/Name"]));
            Assert.ThrowsAny<Exception>(() => sessions.CreateNew(output, ["Sheet1"], "Missing"));
            Assert.ThrowsAny<Exception>(() => sessions.CreateNew(output, ["Bad\0Name"]));
            Assert.False(File.Exists(output));
            Assert.Empty(Directory.GetFiles(directory, "*.staging"));
            var created = JsonSerializer.SerializeToElement(sessions.CreateNew(output));
            Assert.True(File.Exists(output));
            Assert.Throws<IOException>(() => sessions.CreateNew(output));
            Assert.False(JsonSerializer.SerializeToElement(sessions.Status(created.GetProperty("session").GetString()!))
                .GetProperty("source_changed_on_disk").GetBoolean());
            sessions.Close(created.GetProperty("session").GetString()!, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void InvalidOrOccupiedTemplateCreationNeverWritesTheDestination()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-create-errors-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            var destination = Path.Combine(directory, "new.xlsx");
            var invalid = Path.Combine(directory, "invalid.xlsx");
            File.WriteAllText(invalid, "not a ZIP package");
            using var sessions = new ExcelSessions();
            Assert.ThrowsAny<Exception>(() => sessions.CreateFromTemplate(invalid, destination));
            Assert.False(File.Exists(destination));
            Assert.Throws<NotSupportedException>(() => sessions.CreateFromTemplate(source,
                Path.Combine(directory, "new.xlsm")));
            Assert.False(File.Exists(destination));
            Assert.Throws<FileNotFoundException>(() => sessions.CreateFromTemplate(
                Path.Combine(directory, "missing.xlsx"), destination));
            File.WriteAllText(destination, "do not replace");
            Assert.Throws<IOException>(() => sessions.CreateFromTemplate(source, destination));
            Assert.Equal("do not replace", File.ReadAllText(destination));
            Assert.Empty(Directory.GetFiles(directory, "*.staging"));
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Status()).GetProperty("sessions").EnumerateArray());
        }
        finally { Directory.Delete(directory, true); }
    }
}
