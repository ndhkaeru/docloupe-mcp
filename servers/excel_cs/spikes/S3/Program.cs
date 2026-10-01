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
    var result = await client.CallToolAsync("excel_open", new Dictionary<string, object?> { ["path"] = source });
    Console.WriteLine("structured_content=" + result.StructuredContent?.GetRawText());
    Console.WriteLine("text_content=" + (result.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text ?? "<absent>"));
    if (result.IsError == true || result.StructuredContent is null || result.Content.OfType<TextContentBlock>().Count() != 1)
        throw new InvalidOperationException("Dual-channel MCP tool response not received by SDK client");
    var session = result.StructuredContent.Value.GetProperty("data").GetProperty("session").GetString()!;
    var applied = await client.CallToolAsync("excel_apply", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 0, ["sheet"] = "Sheet1",
        ["ops"] = new[] { new { op = "set_value", sheet = "Sheet1", target = "B1", value = 99 } }
    });
    if (applied.IsError == true) throw new InvalidOperationException("Apply failed: " + applied.StructuredContent?.GetRawText());
    var output = Path.Combine(directory, "written.xlsx");
    var saved = await client.CallToolAsync("excel_save", new Dictionary<string, object?>
    {
        ["session"] = session, ["mode"] = "copy", ["path"] = output
    });
    if (saved.IsError == true || saved.StructuredContent?.GetProperty("data").GetProperty("status").GetString() != "verified")
        throw new InvalidOperationException("Verified save failed: " + saved.StructuredContent?.GetRawText());
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
