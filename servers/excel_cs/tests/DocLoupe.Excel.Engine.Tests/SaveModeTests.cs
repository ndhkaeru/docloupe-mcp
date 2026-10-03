using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SaveModeTests
{
    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void SaveAsFollowsNewPathAndOverwriteCanRestoreAnEarlierRevision(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-save-mode-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, variant + ".xlsx");
            var output = Path.Combine(directory, "followed.xlsx");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            var saved = JsonSerializer.SerializeToElement(sessions.Save(id, output, mode: "save_as"));
            Assert.Equal("verified", saved.GetProperty("status").GetString());
            Assert.Equal(1, saved.GetProperty("revision_saved").GetInt32());
            Assert.Equal("27", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            var status = JsonSerializer.SerializeToElement(sessions.Status(id));
            Assert.Equal(output, status.GetProperty("path").GetString());
            Assert.Equal(1, status.GetProperty("saved_revision").GetInt32());
            Assert.False(status.GetProperty("dirty").GetBoolean());

            File.Delete(source);
            Assert.False(JsonSerializer.SerializeToElement(sessions.Status(id))
                .GetProperty("source_changed_on_disk").GetBoolean());
            sessions.Apply(id, 1, [new SetValueOp("Sheet1", "B1", "number", "28")]);
            var copy = Path.Combine(directory, "copy.xlsx");
            sessions.Save(id, copy);
            Assert.Equal("28", Assert.Single(P2aGates.ReadCells(copy, "Sheet1", ["B1"])).Value);
            Assert.Equal("27", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id))
                .GetProperty("saved_revision").GetInt32());

            sessions.Undo(id, 2, 0);
            Assert.Equal("42", Assert.Single(JsonSerializer.SerializeToElement(sessions.Read(id, "Sheet1", ["B1"]))
                .GetProperty("cells").EnumerateArray()).GetProperty("Value").GetString());
            Assert.True(JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("dirty").GetBoolean());
            Assert.Throws<InvalidOperationException>(() => sessions.Close(id, false));
            var overwritten = JsonSerializer.SerializeToElement(sessions.Save(id, null, mode: "overwrite"));
            Assert.Equal(0, overwritten.GetProperty("revision_saved").GetInt32());
            Assert.Equal("42", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["B1"])).Value);
            var backup = overwritten.GetProperty("backup").GetProperty("path").GetString()!;
            Assert.Equal("27", Assert.Single(P2aGates.ReadCells(backup, "Sheet1", ["B1"])).Value);
            Assert.False(JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("dirty").GetBoolean());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ExternalChangesToFollowedOutputBlockFurtherEdits()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-save-mode-stale-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            var output = Path.Combine(directory, "followed.xlsx");
            sessions.Save(id, output, mode: "save_as");
            File.WriteAllText(output, "modified externally");
            Assert.True(JsonSerializer.SerializeToElement(sessions.Status(id))
                .GetProperty("source_changed_on_disk").GetBoolean());
            Assert.Throws<InvalidOperationException>(() => sessions.Apply(id, 1,
                [new SetValueOp("Sheet1", "C2", "number", "7")]));
            Assert.Throws<InvalidOperationException>(() => sessions.Save(id, null, mode: "overwrite"));
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id))
                .GetProperty("saved_revision").GetInt32());
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void FailedSaveAsAndOverwriteLeaveOutputAndRevisionUntouched()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-save-mode-blocked-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            var followed = Path.Combine(directory, "followed.xlsx");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            var impossible = new ValueAssertion("Sheet1", "B1", true, "number", "999", null);
            Assert.Throws<SaveBlockedException>(() => sessions.Save(id, followed, [impossible], "save_as"));
            Assert.False(File.Exists(followed));
            Assert.Equal(source, JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("path").GetString());
            sessions.Save(id, followed, mode: "save_as");
            sessions.Apply(id, 1, [new SetValueOp("Sheet1", "B1", "number", "28")]);
            var before = File.ReadAllBytes(followed);
            Assert.Throws<SaveBlockedException>(() => sessions.Save(id, null, [impossible], "overwrite"));
            Assert.Equal(before, File.ReadAllBytes(followed));
            Assert.Equal(1, JsonSerializer.SerializeToElement(sessions.Status(id))
                .GetProperty("saved_revision").GetInt32());
            Assert.Empty(Directory.GetFiles(directory, "*.bak"));
            Assert.Throws<ArgumentException>(() => sessions.Save(id, source, mode: "overwrite"));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }
}
