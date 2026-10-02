using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.AI;
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
    McpServerTool.Create((string path, string detail = "summary", string? sheet = null, int max_rows = 20, int max_cols = 10) =>
        Handle(() => sessions.Peek(path, detail, sheet, max_rows, max_cols)), new McpServerToolCreateOptions { Name = "excel_peek" }),
    McpServerTool.Create((string after_path) => Handle(() => sessions.Verify(after_path)),
        new McpServerToolCreateOptions { Name = "excel_verify" }),
    McpServerTool.Create((string? session = null) => Handle(() => sessions.Status(session)), new McpServerToolCreateOptions { Name = "excel_status" }),
    McpServerTool.Create((string session, string? sheet, JsonElement target) => Handle(() => sessions.Read(session, sheet, ReadTargets(target))),
        new McpServerToolCreateOptions
        {
            Name = "excel_read",
            SchemaCreateOptions = new AIJsonSchemaCreateOptions
            {
                TransformSchemaNode = (context, node) => context.TypeInfo.Type == typeof(JsonElement)
                    ? new JsonObject
                    {
                        ["oneOf"] = new JsonArray(
                            new JsonObject { ["type"] = "string" },
                            new JsonObject
                            {
                                ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" },
                                ["minItems"] = 1, ["maxItems"] = 500
                            })
                    }
                    : node
            }
        }),
    McpServerTool.Create((string session, int base_revision, SetValueRequest[] ops, string? sheet) => Handle(() => sessions.Apply(session, base_revision, ops.SelectMany((op, index) => op.NormalizeMany(sheet).Select(cell => cell with { SourceIndex = index })).ToArray())), new McpServerToolCreateOptions { Name = "excel_apply" }),
    McpServerTool.Create((string session, string mode, string path, SaveAssertionRequest[]? @assert = null) => Handle(() => mode == "copy" ? sessions.Save(session, path, @assert?.Select(item => item.Normalize()).ToArray()) : throw new NotSupportedException("P2a save supports copy mode only")), new McpServerToolCreateOptions { Name = "excel_save" }),
    McpServerTool.Create((string session, bool discard_unsaved) => Handle(() => sessions.Close(session, discard_unsaved)), new McpServerToolCreateOptions { Name = "excel_close" }),
    McpServerTool.Create((string session, int base_revision, int to_revision) => Handle(() =>
    {
        var undone = sessions.Undo(session, base_revision, to_revision);
        return new { revision = undone.Revision, discarded = undone.Discarded };
    }), new McpServerToolCreateOptions { Name = "excel_undo" })
]);
try { await builder.Build().RunAsync(); }
finally { sessions.Dispose(); }

static string[] ReadTargets(JsonElement target)
{
    if (target.ValueKind == JsonValueKind.String)
        return [target.GetString() ?? throw new ArgumentException("Target cannot be null")];
    if (target.ValueKind != JsonValueKind.Array || target.GetArrayLength() is < 1 or > 500)
        throw new ArgumentException("Target must be a cell, range, or nonempty array of at most 500 targets");
    return target.EnumerateArray().Select(item => item.ValueKind == JsonValueKind.String
        ? item.GetString()! : throw new ArgumentException("Every read target must be a string")).ToArray();
}

static CallToolResult Handle(Func<object> action)
{
    try
    {
        var data = action();
        if (data is ReadOnlyVerification verification)
        {
            var summary = verification.Summary;
            var report = new { mode = "validate", files = new { after = verification.Path }, status = summary.Status, partial = true,
                package_issues = summary.PackageIssues, markup_issues = summary.MarkupIssues,
                unverified_gates = summary.UnverifiedGates };
            if (summary.Status == "failed")
                return Result(new { ok = false, error = new { code = "PACKAGE_INVALID",
                    message = "Read-only verification found package or markup issues", details = report,
                    retryable = false } }, true);
            return Result(new { ok = true, data = report, warnings = new[] { "Partial verification only; unverified gates remain" } }, false);
        }
        return Result(new { ok = true, data, warnings = Array.Empty<object>() }, false);
    }
    catch (SaveBlockedException blocked)
    {
        return Result(new { ok = false, error = new { code = "SAVE_BLOCKED", message = "Staging verification failed; no output was written",
            details = blocked.Issues, retryable = false } }, true);
    }
    catch (PreconditionFailedException failed)
    {
        return Result(new { ok = false, error = new { code = "PRECONDITION_FAILED", message = failed.Message,
            details = new { index = failed.Index, target = failed.Target, expected = failed.Expected,
                actual = failed.Actual }, retryable = false } }, true);
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
