using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.AI;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
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
    McpServerTool.Create((string target_path, string? template_path = null, string[]? sheets = null,
        string? active_sheet = null, string? format = null,
        CreatePropertiesRequest? document_properties = null) => Handle(() =>
        {
            if (format is not null && !Path.GetExtension(target_path).Equals("." + format, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("format must match target_path extension");
            if (template_path is not null)
            {
                if (sheets is not null || active_sheet is not null || document_properties is not null)
                    throw new NotSupportedException("Template creation does not support sheet or document-property changes");
                return sessions.CreateFromTemplate(template_path, target_path);
            }
            return sessions.CreateNew(target_path, sheets, active_sheet, document_properties?.Normalize());
        }), new McpServerToolCreateOptions { Name = "excel_create" }),
    McpServerTool.Create((string path, string detail = "summary", string? sheet = null, int max_rows = 20, int max_cols = 10) =>
        Handle(() => sessions.Peek(path, detail, sheet, max_rows, max_cols)), new McpServerToolCreateOptions { Name = "excel_peek" }),
    McpServerTool.Create((string after_path, string? before_path = null, int max_differences = 200, SaveAssertionRequest[]? @assert = null) =>
        Handle(() => before_path is null
            ? sessions.Verify(after_path, NormalizeAssertions(@assert))
            : sessions.Verify(after_path, before_path, max_differences, NormalizeAssertions(@assert))),
        new McpServerToolCreateOptions { Name = "excel_verify" }),
    McpServerTool.Create((string? session = null) => Handle(() => sessions.Status(session)), new McpServerToolCreateOptions { Name = "excel_status" }),
    McpServerTool.Create((string session, JsonElement? target = null, string? sheet = null, bool skip_empty = true, string view = "cells", string[]? include = null) =>
        Handle(() => sessions.Read(session, sheet, ReadTargets(target), skip_empty, view, include)),
        new McpServerToolCreateOptions
        {
            Name = "excel_read",
            SchemaCreateOptions = new AIJsonSchemaCreateOptions
            {
                TransformSchemaNode = (context, node) => context.TypeInfo.Type == typeof(JsonElement) || context.TypeInfo.Type == typeof(JsonElement?)
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
    McpServerTool.Create((string session, FindQueryRequest query, FindScopeRequest? scope = null, string @in = "value", bool case_sensitive = false, string normalize = "nfc", int max_results = 100) =>
        Handle(() =>
        {
            if (scope?.Other is { Count: > 0 }) throw new NotSupportedException("Unsupported find scope fields");
            var (pattern, isRegex, expectedValue, formulaContains) = query.Normalize();
            if (formulaContains && @in is not ("value" or "formula"))
                throw new NotSupportedException("formula_contains only searches formula text");
            return sessions.Find(session, scope?.Sheet, scope?.Target, pattern, isRegex,
                formulaContains ? "formula" : @in, case_sensitive, normalize,
                max_results, expectedValue);
        }), new McpServerToolCreateOptions { Name = "excel_find" }),
    McpServerTool.Create((string session, int base_revision, SetValueRequest[] ops, string? sheet = null, bool dry_run = false, int max_diff_items = 200, string @return = "diff+readback") => Handle(() => sessions.Apply(session, base_revision, ops.SelectMany((op, index) => op.NormalizeMany(sheet).Select(cell => cell with { SourceIndex = index })).ToArray(), dry_run, max_diff_items, @return)), new McpServerToolCreateOptions { Name = "excel_apply" }),
    McpServerTool.Create((string session, string mode, string? path = null, SaveAssertionRequest[]? @assert = null) =>
        Handle(() => sessions.Save(session, path, NormalizeAssertions(@assert), mode)),
        new McpServerToolCreateOptions { Name = "excel_save" }),
    McpServerTool.Create((string session, bool discard_unsaved = false) => Handle(() => sessions.Close(session, discard_unsaved)), new McpServerToolCreateOptions { Name = "excel_close" }),
    McpServerTool.Create((string session, int base_revision, int to_revision) => Handle(() =>
    {
        var undone = sessions.Undo(session, base_revision, to_revision);
        return new { revision = undone.Revision, discarded = undone.Discarded };
    }), new McpServerToolCreateOptions { Name = "excel_undo" })
]);
try { await builder.Build().RunAsync(); }
finally { sessions.Dispose(); }

static ValueAssertion[]? NormalizeAssertions(SaveAssertionRequest[]? requests)
{
    if (requests is null) return null;
    if (requests.Length > 500) throw new ArgumentException("At most 500 assertions are supported");
    return requests.Select(item => item.Normalize()).ToArray();
}

static string[] ReadTargets(JsonElement? target)
{
    if (target is null || target.Value.ValueKind == JsonValueKind.Undefined) return [];
    if (target.Value.ValueKind == JsonValueKind.String)
        return [target.Value.GetString() ?? throw new ArgumentException("Target cannot be null")];
    if (target.Value.ValueKind != JsonValueKind.Array || target.Value.GetArrayLength() is < 1 or > 500)
        throw new ArgumentException("Target must be a cell, range, or nonempty array of at most 500 targets");
    return target.Value.EnumerateArray().Select(item => item.ValueKind == JsonValueKind.String
        ? item.GetString()! : throw new ArgumentException("Every read target must be a string")).ToArray();
}

static CallToolResult Handle(Func<object> action)
{
    try
    {
        var data = action();
        if (data is ReadOnlyComparison comparison)
        {
            var report = new
            {
                mode = "compare", files = new { before = comparison.Before.Path, after = comparison.After.Path },
                status = comparison.Status, partial = true,
                checks_run = comparison.Comparison is null ? Array.Empty<string>() :
                    comparison.AssertionIssues is null ? ["G5_PART_BYTES"] : ["G5_PART_BYTES", "G7_ASSERTIONS"],
                before = new { package_issues = comparison.Before.Summary.PackageIssues,
                    markup_issues = comparison.Before.Summary.MarkupIssues,
                    schema_issues = comparison.Before.Schema?.Issues ?? [], schema_gaps = comparison.Before.Schema?.Gaps ?? [] },
                after = new { package_issues = comparison.After.Summary.PackageIssues,
                    markup_issues = comparison.After.Summary.MarkupIssues,
                    schema_issues = comparison.After.Schema?.Issues ?? [], schema_gaps = comparison.After.Schema?.Gaps ?? [] },
                differences = comparison.Comparison?.Differences ?? [],
                new_schema_issues = comparison.SchemaDelta?.Issues ?? [],
                schema_gaps = comparison.SchemaDelta?.Gaps ?? [],
                assertion_issues = comparison.AssertionIssues ?? [],
                truncated = comparison.Comparison?.Truncated ?? false,
                unverified_gates = new[] { "G1_REMAINING", "G2", "G3_REMAINING", "G4", "G5_REMAINING", "G6", "G7" }
            };
            if (comparison.Status == "failed")
            {
                var code = comparison.Comparison is null ? "PACKAGE_INVALID"
                    : comparison.Comparison.HasDifferences || comparison.SchemaDelta?.Issues.Count > 0
                        ? "PRESERVATION_FAILED" : "ASSERTION_FAILED";
                return Result(new { ok = false, error = new { code, message = "Read-only comparison failed",
                    details = report, retryable = false } }, true);
            }
            return Result(new { ok = true, data = report,
                warnings = new[] { "Part bytes only; ZIP metadata and semantic equivalence are not verified" } }, false);
        }
        if (data is ReadOnlyVerification verification)
        {
            var summary = verification.Summary;
            var report = new { mode = "validate", files = new { after = verification.Path }, status = summary.Status, partial = true,
                package_issues = summary.PackageIssues, markup_issues = summary.MarkupIssues,
                schema_issues = verification.Schema?.Issues ?? [], schema_gaps = verification.Schema?.Gaps ?? [],
                assertion_issues = verification.AssertionIssues ?? [],
                unverified_gates = summary.UnverifiedGates };
            if (summary.Status == "failed")
                return Result(new { ok = false, error = new {
                    code = verification.AssertionIssues is { Count: > 0 } ? "ASSERTION_FAILED" : "PACKAGE_INVALID",
                    message = "Read-only verification failed", details = report,
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
        var details = new Dictionary<string, object?>
        {
            ["index"] = failed.Index, ["target"] = failed.Target,
            ["expected"] = failed.Expected, ["actual"] = failed.Actual
        };
        if (failed.Expected.FontBold is not null || failed.Expected.FontItalic is not null)
        {
            var font = new Dictionary<string, bool?>();
            if (failed.Expected.FontBold is not null) font["bold"] = failed.ActualFontBold;
            if (failed.Expected.FontItalic is not null) font["italic"] = failed.ActualFontItalic;
            details["actual_style"] = font.Values.All(value => value is null) ? null : new { font };
        }
        return Result(new { ok = false, error = new { code = "PRECONDITION_FAILED", message = failed.Message,
            details, retryable = false } }, true);
    }
    catch (Exception error)
    {
        var mapped = ToolErrors.FromException(error);
        return Result(new { ok = false, error = new { code = mapped.Code, message = mapped.Message,
            retryable = mapped.Retryable } }, true);
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
