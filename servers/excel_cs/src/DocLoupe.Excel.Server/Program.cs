using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

var sessions = new ExcelSessions();
var builder = Host.CreateApplicationBuilder(args);
builder.Logging.ClearProviders();
builder.Services.AddMcpServer().WithStdioServerTransport().WithTools([
    McpServerTool.Create((string path) => Handle(() => sessions.Open(path)), new McpServerToolCreateOptions { Name = "excel_open" }),
    McpServerTool.Create((string session, string sheet, string target) => Handle(() => sessions.Read(session, sheet, [target])), new McpServerToolCreateOptions { Name = "excel_read" }),
    McpServerTool.Create((string session, int base_revision, SetValueRequest[] ops, string? sheet) => Handle(() => sessions.Apply(session, base_revision, ops.Select(op => op.Normalize(sheet)).ToArray())), new McpServerToolCreateOptions { Name = "excel_apply" }),
    McpServerTool.Create((string session, string mode, string path) => Handle(() => mode == "copy" ? sessions.Save(session, path) : throw new NotSupportedException("P2a save supports copy mode only")), new McpServerToolCreateOptions { Name = "excel_save" }),
    McpServerTool.Create((string session, bool discard_unsaved) => Handle(() => sessions.Close(session, discard_unsaved)), new McpServerToolCreateOptions { Name = "excel_close" })
]);
try { await builder.Build().RunAsync(); }
finally { sessions.Dispose(); }

static CallToolResult Handle(Func<object> action)
{
    try { return Result(new { ok = true, data = action(), warnings = Array.Empty<object>() }, false); }
    catch (SaveBlockedException blocked)
    {
        return Result(new { ok = false, error = new { code = "SAVE_BLOCKED", message = "Staging verification failed; no output was written",
            details = blocked.Issues, retryable = false } }, true);
    }
    catch (Exception error)
    {
        return Result(new { ok = false, error = new { code = "P2A_ERROR", message = error.Message, retryable = false } }, true);
    }
}

static CallToolResult Result<T>(T envelope, bool error)
{
    var json = JsonSerializer.SerializeToElement(envelope);
    return new CallToolResult
    {
        IsError = error,
        StructuredContent = json,
        Content = [new TextContentBlock { Text = json.GetRawText() }]
    };
}
