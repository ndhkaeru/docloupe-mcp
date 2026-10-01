using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class StatusTests
{
    [Fact]
    public void ListsSessionsAndTracksRevisionsAcrossCopySaveAndUndo()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-status-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var opened = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")));
            var id = opened.GetProperty("session").GetString()!;
            var initial = JsonSerializer.SerializeToElement(sessions.Status(id));
            Assert.Equal(0, initial.GetProperty("revision").GetInt32());
            Assert.False(initial.GetProperty("dirty").GetBoolean());
            Assert.False(initial.GetProperty("source_changed_on_disk").GetBoolean());
            Assert.Empty(initial.GetProperty("ledger").EnumerateArray());
            Assert.Equal("unavailable", initial.GetProperty("server").GetProperty("oracles").GetProperty("excel").GetString());

            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            sessions.Apply(id, 1, [new SetValueOp("Sheet1", "B1", "number", "28"),
                new SetValueOp("Sheet1", "D3", "number", "8")]);
            var listing = JsonSerializer.SerializeToElement(sessions.Status());
            var listed = Assert.Single(listing.GetProperty("sessions").EnumerateArray());
            Assert.Equal(id, listed.GetProperty("session").GetString());
            Assert.Equal(2, listed.GetProperty("revision").GetInt32());
            Assert.True(listed.GetProperty("dirty").GetBoolean());
            var state = JsonSerializer.SerializeToElement(sessions.Status(id));
            Assert.Equal(0, state.GetProperty("saved_revision").GetInt32());
            Assert.Equal(2, state.GetProperty("ledger")[1].GetProperty("op_count").GetInt32());
            Assert.Contains("set_value Sheet1!D3", state.GetProperty("ledger")[1].GetProperty("summary").GetString());

            sessions.Save(id, Path.Combine(directory, "copy.xlsx"));
            Assert.True(JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("dirty").GetBoolean());
            sessions.Undo(id, 2, 1);
            state = JsonSerializer.SerializeToElement(sessions.Status(id));
            Assert.Equal(1, state.GetProperty("revision").GetInt32());
            Assert.Single(state.GetProperty("ledger").EnumerateArray());
            sessions.Undo(id, 1, 0);
            Assert.False(JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("dirty").GetBoolean());
            sessions.Close(id, false);
            Assert.Empty(JsonSerializer.SerializeToElement(sessions.Status()).GetProperty("sessions").EnumerateArray());
            Assert.Throws<KeyNotFoundException>(() => sessions.Status(id));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void LedgerOnlyExposesLastTwentyRevisions()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-status-ledger-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            for (var revision = 0; revision < 22; revision++)
                sessions.Apply(id, revision, [new SetValueOp("Sheet1", "B1", "number", (revision + 1).ToString())]);
            var ledger = JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("ledger");
            Assert.Equal(20, ledger.GetArrayLength());
            Assert.Equal(3, ledger[0].GetProperty("revision").GetInt32());
            Assert.Equal(22, ledger[19].GetProperty("revision").GetInt32());
            sessions.Undo(id, 22, 1);
            ledger = JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("ledger");
            Assert.Equal(1, Assert.Single(ledger.EnumerateArray()).GetProperty("revision").GetInt32());
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ReportsSourceChangeWithoutBlockingStatus()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-status-change-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var path = Path.Combine(directory, "default.xlsx");
            var id = JsonSerializer.SerializeToElement(sessions.Open(path)).GetProperty("session").GetString()!;
            File.AppendAllText(path, "changed");
            Assert.True(JsonSerializer.SerializeToElement(sessions.Status(id)).GetProperty("source_changed_on_disk").GetBoolean());
            Assert.Throws<InvalidOperationException>(() => sessions.Read(id, "Sheet1", ["B1"]));
            sessions.Close(id, false);
        }
        finally { Directory.Delete(directory, true); }
    }
}
