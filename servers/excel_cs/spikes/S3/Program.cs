using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

if (args is not [var serverDll]) throw new ArgumentException("Usage: <server-dll-path>");
var directory = Path.Combine(Path.GetTempPath(), "docloupe-s3-" + Guid.NewGuid().ToString("N"));
try
{
    var source = SyntheticFixtures.Create(directory)[0];
    var transport = new StdioClientTransport(new StdioClientTransportOptions
    {
        Command = "dotnet",
        Arguments = [Path.GetFullPath(serverDll)],
        Name = "p2a-s3"
    });
    await using var client = await McpClient.CreateAsync(transport);
    var tools = await client.ListToolsAsync();
    Console.WriteLine("tools=" + string.Join(',', tools.Select(tool => tool.Name)));
    if (!tools.Any(tool => tool.Name == "excel_undo")) throw new InvalidOperationException("Undo tool is missing");
    if (!tools.Any(tool => tool.Name == "excel_status")) throw new InvalidOperationException("Status tool is missing");
    var emptyStatus = await client.CallToolAsync("excel_status", new Dictionary<string, object?>());
    if (emptyStatus.IsError == true || emptyStatus.StructuredContent?.GetProperty("data").GetProperty("sessions").GetArrayLength() != 0)
        throw new InvalidOperationException("Empty status failed: " + emptyStatus.StructuredContent?.GetRawText()
            + " text=" + emptyStatus.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text
            + " schema=" + tools.Single(tool => tool.Name == "excel_status").ProtocolTool.InputSchema.GetRawText());
    var result = await client.CallToolAsync("excel_open", new Dictionary<string, object?> { ["path"] = source });
    Console.WriteLine("structured_content=" + result.StructuredContent?.GetRawText());
    Console.WriteLine("text_content=" + (result.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text ?? "<absent>"));
    if (result.IsError == true || result.StructuredContent is null || result.Content.OfType<TextContentBlock>().Count() != 1)
        throw new InvalidOperationException("Dual-channel MCP tool response not received by SDK client");
    var session = result.StructuredContent.Value.GetProperty("data").GetProperty("session").GetString()!;
    var applied = await client.CallToolAsync("excel_apply", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 0, ["sheet"] = "Sheet1",
        ["ops"] = new object[] { new { op = "set_value", sheet = "Sheet1", target = "B1", value = 99 },
            new { op = "set_formula", sheet = "Sheet1", target = "C1", formula = "=B1+2", cache = "clear" },
            new { op = "set_values", sheet = "Sheet1", target = "D4", values = new object?[][] { [4, "batch"] } },
            new { op = "set_value", sheet = "Sheet1", target = "F5:G5", value = 6 },
            new { op = "fill", sheet = "Sheet1", target = "H6:I6", value = 8 },
            new { op = "fill", sheet = "Sheet1", target = "J7:K7", series = new { start = -3, step = 2 } },
            new { op = "clear", sheet = "Sheet1", target = "A1", what = new[] { "values" }, remove_cells = false },
            new { op = "clear", sheet = "Sheet1", target = "L8" },
            new { op = "clear", sheet = "Sheet1", target = "D3", remove_cells = true },
            new { op = "set_value", sheet = "Sheet1", target = "M9", value = new { error = "#N/A" } },
            new { op = "fill", sheet = "Sheet1", target = "N10:O10", series = new { start = 0.1, step = 0.2 } } }
    });
    if (applied.IsError == true) throw new InvalidOperationException("Apply failed: " + applied.StructuredContent?.GetRawText());
    var status = await client.CallToolAsync("excel_status", new Dictionary<string, object?> { ["session"] = session });
    if (status.IsError == true || status.StructuredContent?.GetProperty("data").GetProperty("revision").GetInt32() != 1 ||
        status.StructuredContent?.GetProperty("data").GetProperty("ledger").GetArrayLength() != 1)
        throw new InvalidOperationException("Session status failed: " + status.StructuredContent?.GetRawText());
    var invalidOutput = Path.Combine(directory, "invalid.xlsx");
    var rejected = await client.CallToolAsync("excel_save", new Dictionary<string, object?>
    {
        ["session"] = session, ["mode"] = "copy", ["path"] = invalidOutput,
        ["assert"] = new[] { new { target = "Sheet1!B1", equals = new { display = "99" } } }
    });
    if (rejected.IsError != true || File.Exists(invalidOutput))
        throw new InvalidOperationException("Unsupported assertion was not rejected");
    var output = Path.Combine(directory, "written.xlsx");
    var saved = await client.CallToolAsync("excel_save", new Dictionary<string, object?>
    {
        ["session"] = session, ["mode"] = "copy", ["path"] = output,
        ["assert"] = new object[] { new { target = "Sheet1!B1", equals = new { value = 99 } },
            new { target = "Sheet1!C1", equals = new { formula = "B1+2" } },
            new { target = "Sheet1!D4", equals = new { value = 4 } },
            new { target = "Sheet1!E4", equals = new { value = "batch" } },
            new { target = "Sheet1!F5", equals = new { value = 6 } },
            new { target = "Sheet1!G5", equals = new { value = 6 } },
            new { target = "Sheet1!H6", equals = new { value = 8 } },
            new { target = "Sheet1!I6", equals = new { value = 8 } },
            new { target = "Sheet1!J7", equals = new { value = -3 } },
            new { target = "Sheet1!K7", equals = new { value = -1 } },
            new { target = "Sheet1!A1", equals = new { value = System.Text.Json.JsonSerializer.SerializeToElement<object?>(null) } },
            new { target = "Sheet1!L8", equals = new { value = System.Text.Json.JsonSerializer.SerializeToElement<object?>(null) } },
            new { target = "Sheet1!D3", equals = new { value = System.Text.Json.JsonSerializer.SerializeToElement<object?>(null) } },
            new { target = "Sheet1!M9", equals = new { value = new { error = "#N/A" } } },
            new { target = "Sheet1!N10", equals = new { value = 0.1 } },
            new { target = "Sheet1!O10", equals = new { value = 0.3 } } }
    });
    if (saved.IsError == true || saved.StructuredContent?.GetProperty("data").GetProperty("status").GetString() != "verified")
        throw new InvalidOperationException("Verified save failed: " + saved.StructuredContent?.GetRawText());
    var verifiedGates = saved.StructuredContent.Value.GetProperty("data").GetProperty("gates").EnumerateArray()
        .Select(gate => gate.GetString()).ToArray();
    if (!verifiedGates.Contains("G6") || !verifiedGates.Contains("G7"))
        throw new InvalidOperationException("G6 or G7 was not checked");
    Console.WriteLine("saved_status=" + saved.StructuredContent.Value.GetProperty("data").GetProperty("status").GetString());
    var undone = await client.CallToolAsync("excel_undo", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 1, ["to_revision"] = 0
    });
    if (undone.IsError == true || undone.StructuredContent?.GetProperty("data").GetProperty("revision").GetInt32() != 0)
        throw new InvalidOperationException("Undo failed: " + undone.StructuredContent?.GetRawText());
    var closed = await client.CallToolAsync("excel_close", new Dictionary<string, object?>
    {
        ["session"] = session, ["discard_unsaved"] = false
    });
    if (closed.IsError == true) throw new InvalidOperationException("Close failed");
}
finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
