using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

var builder = Host.CreateApplicationBuilder(args);
builder.Logging.ClearProviders();
builder.Services.AddMcpServer().WithStdioServerTransport().WithTools([
    McpServerTool.Create(() => new CallToolResult
    {
        StructuredContent = JsonSerializer.SerializeToElement(new { channel = "STRUCTURED_ONLY_P2A_73" }),
        Content = [new TextContentBlock { Text = "{\"channel\":\"TEXT_ONLY_P2A_91\"}" }]
    }, new McpServerToolCreateOptions { Name = "channel_probe", Description = "Returns two different channel markers for client forwarding measurement" })
]);
await builder.Build().RunAsync();
