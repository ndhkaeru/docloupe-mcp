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
