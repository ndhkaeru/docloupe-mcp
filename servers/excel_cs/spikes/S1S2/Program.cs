using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using ModelContextProtocol.Protocol;

if (args is ["probe", var sourceRoot, var outputRoot])
{
    Probe.Run(sourceRoot, outputRoot);
    return;
}

if (args is ["probe-autosave", var autoSaveSourceRoot, var autoSaveOutputRoot])
{
    Probe.RunAutoSaveCheck(autoSaveSourceRoot, autoSaveOutputRoot);
    return;
}

var builder = Host.CreateApplicationBuilder(args);
builder.Logging.ClearProviders();
builder.Services.AddMcpServer().WithStdioServerTransport().WithTools(
    [McpServerTool.Create((string path) =>
    {
        var result = Probe.ReadFirstSheet(path);
        return new CallToolResult
        {
            Content = [new TextContentBlock { Text = result.FirstCell }],
            StructuredContent = JsonSerializer.SerializeToElement(result, ProbeJsonContext.Default.SheetResult)
        };
    },
        new McpServerToolCreateOptions { Name = "read_first_sheet", SerializerOptions = new JsonSerializerOptions { TypeInfoResolver = ProbeJsonContext.Default } })]);
await builder.Build().RunAsync();

internal sealed record SheetResult(string Name, string FirstCell);

[JsonSerializable(typeof(SheetResult))]
[JsonSerializable(typeof(CallToolResult))]
[JsonSerializable(typeof(TextContentBlock))]
internal partial class ProbeJsonContext : JsonSerializerContext;

internal static class Probe
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace X14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private static readonly XNamespace Unknown = "urn:docloupe:spike:unknown";
    private const string Marker = "S1S2-edited";

    public static SheetResult ReadFirstSheet(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var firstSheet = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().First();
        var worksheet = ((WorksheetPart)document.WorkbookPart.GetPartById(firstSheet.Id!.Value!)).Worksheet!;
        return new SheetResult(firstSheet.Name!.Value!, worksheet.Descendants<Cell>().FirstOrDefault()?.InnerText ?? "");
    }

    private static List<string> ListFixtures(string sourceRoot)
    {
        var fixtures = Directory.GetFiles(sourceRoot, "*", SearchOption.TopDirectoryOnly)
            .Where(path => Path.GetFileName(path).Length > 2 && Path.GetFileName(path)[..2] is "00" or "01" or "02" or "03" or "04" or "05" or "06" or "07")
            .Where(path => !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal))
            .OrderBy(path => path).ToList();
        if (fixtures.Count != 8) throw new InvalidOperationException($"Expected 8 fixtures, found {fixtures.Count}");
        return fixtures;
    }

    // Approach A again, with AutoSave on and off. Shows whether the xl/workbook.xml rewrite seen
    // in `probe` comes from AutoSave saving the workbook DOM that FirstPart has to load.
    public static void RunAutoSaveCheck(string sourceRoot, string outputRoot)
    {
        Directory.CreateDirectory(outputRoot);
        foreach (var source in ListFixtures(sourceRoot))
        foreach (var autoSave in new[] { true, false })
        {
            var output = Path.Combine(outputRoot, Path.GetFileNameWithoutExtension(source) + $"_A_autosave-{autoSave}" + Path.GetExtension(source));
            try
            {
                File.Copy(source, output, true);
                string part;
                using (var document = SpreadsheetDocument.Open(output, true, new OpenSettings { AutoSave = autoSave }))
                {
                    var (worksheetPart, name) = FirstPart(document);
                    part = name;
                    worksheetPart.Worksheet!.Save();
                }
                using var beforeZip = ZipFile.OpenRead(source);
                using var afterZip = ZipFile.OpenRead(output);
                var changed = beforeZip.Entries.Where(entry => entry.FullName != part)
                    .Where(entry => afterZip.GetEntry(entry.FullName) is not { } after || !Read(entry).AsSpan().SequenceEqual(Read(after)))
                    .Select(entry => entry.FullName).ToList();
                Console.WriteLine($"{Path.GetFileName(source)} | A | autosave={autoSave} | other_bytes={(changed.Count == 0 ? "same" : string.Join(',', changed))}");
            }
            catch (Exception exception)
            {
                Console.WriteLine($"{Path.GetFileName(source)} | A | autosave={autoSave} | ERROR {exception.GetType().Name}: {exception.Message.Replace('\n', ' ')}");
            }
        }
    }

    public static void Run(string sourceRoot, string outputRoot)
    {
        Directory.CreateDirectory(outputRoot);
        var fixtures = ListFixtures(sourceRoot);
        var syntheticSource = Path.Combine(outputRoot, "V06_EX03_valid.xlsx");
        CreateSynthetic(fixtures[0], syntheticSource, false);
        fixtures.Add(syntheticSource);
        var invalidSource = Path.Combine(outputRoot, "EX04_invalid.xlsx");
        CreateSynthetic(fixtures[0], invalidSource, true);
        foreach (var source in fixtures.Append(invalidSource))
        {
            foreach (var approach in new[] { "A", "B" })
            foreach (var edit in new[] { false, true })
            {
                var output = Path.Combine(outputRoot, Path.GetFileNameWithoutExtension(source) + $"_{approach}_{(edit ? "edit" : "noop")}" + Path.GetExtension(source));
                try
                {
                    string part = approach == "A" ? WriteA(source, output, edit) : WriteB(source, output, edit);
                    var result = Compare(source, output, part, edit);
                    Console.WriteLine($"{Path.GetFileName(source)} | {approach} | {(edit ? "edit" : "noop")} | {result}");
                }
                catch (Exception exception)
                {
                    Console.WriteLine($"{Path.GetFileName(source)} | {approach} | {(edit ? "edit" : "noop")} | ERROR {exception.GetType().Name}: {exception.Message.Replace('\n', ' ')}");
                }
            }
        }
    }

    private static (WorksheetPart Part, string Name) FirstPart(SpreadsheetDocument document)
    {
        var workbook = document.WorkbookPart!;
        var firstSheet = workbook.Workbook!.Sheets!.Elements<Sheet>().First();
        var part = (WorksheetPart)workbook.GetPartById(firstSheet.Id!.Value!);
        return (part, part.Uri.ToString().TrimStart('/'));
    }

    private static string WriteA(string source, string output, bool edit)
    {
        File.Copy(source, output, true);
        using var document = SpreadsheetDocument.Open(output, true);
        var (part, name) = FirstPart(document);
        if (edit) Edit(part.Worksheet!);
        part.Worksheet!.Save();
        return name;
    }

    private static string WriteB(string source, string output, bool edit)
    {
        string name;
        string xml;
        using var normalizedPackage = NormalizeContentTypesForSdk(source);
        using (var document = SpreadsheetDocument.Open(normalizedPackage, false))
        {
            var (part, partName) = FirstPart(document);
            name = partName;
            var detached = (Worksheet)part.Worksheet!.CloneNode(true);
            if (edit) Edit(detached);
            xml = detached.OuterXml;
        }
        using var input = ZipFile.OpenRead(source);
        using var destination = new ZipArchive(File.Create(output), ZipArchiveMode.Create);
        foreach (var entry in input.Entries)
        {
            var target = destination.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            target.LastWriteTime = entry.LastWriteTime;
            using var targetStream = target.Open();
            if (entry.FullName == name)
            {
                using var writer = new StreamWriter(targetStream, new UTF8Encoding(false), leaveOpen: true);
                writer.Write(xml);
            }
            else
            {
                using var sourceStream = entry.Open();
                sourceStream.CopyTo(targetStream);
            }
        }
        return name;
    }

    private static void Edit(Worksheet worksheet)
    {
        var cell = worksheet.Descendants<Cell>().First();
        cell.DataType = CellValues.InlineString;
        cell.CellValue = null;
        cell.InlineString = new InlineString(new Text(Marker));
    }

    private static string Compare(string source, string output, string part, bool edit)
    {
        using var beforeZip = ZipFile.OpenRead(source);
        using var afterZip = ZipFile.OpenRead(output);
        var mismatches = beforeZip.Entries.Where(entry => entry.FullName != part)
            .Where(entry => afterZip.GetEntry(entry.FullName) is not { } after || !Read(entry).AsSpan().SequenceEqual(Read(after)))
            .Select(entry => entry.FullName).ToList();
        mismatches.AddRange(afterZip.Entries.Where(entry => beforeZip.GetEntry(entry.FullName) is null).Select(entry => "+" + entry.FullName));
        var oldXml = XDocument.Parse(Encoding.UTF8.GetString(Read(beforeZip.GetEntry(part)!)));
        var newBytes = Read(afterZip.GetEntry(part)!);
        var newXml = XDocument.Parse(Encoding.UTF8.GetString(newBytes));
        var oldRoot = oldXml.Root!;
        var newRoot = newXml.Root!;
        var oldCell = oldRoot.Descendants(Main + "c").First();
        var newCell = newRoot.Descendants(Main + "c").First();
        var cellRef = (string?)oldCell.Attribute("r");
        var cellValue = newCell.Descendants(Main + "t").FirstOrDefault()?.Value;
        var observed = edit ? cellValue == Marker : CellSignature(oldCell) == CellSignature(newCell);
        var oldSignature = ElementSignature(oldRoot, cellRef);
        var newSignature = ElementSignature(newRoot, cellRef);
        var preserved = oldSignature == newSignature;
        var oldNs = oldRoot.Attributes().Where(attribute => attribute.IsNamespaceDeclaration)
            .ToDictionary(attribute => attribute.Name.LocalName, attribute => attribute.Value);
        var newNs = newRoot.Attributes().Where(attribute => attribute.IsNamespaceDeclaration)
            .ToDictionary(attribute => attribute.Name.LocalName, attribute => attribute.Value);
        var nsPreserved = oldNs.All(pair => newNs.TryGetValue(pair.Key, out var value) && value == pair.Value);
        var ignorable = (string?)newRoot.Attribute(Mc + "Ignorable") ?? "";
        var mcValid = ignorable.Split(' ', StringSplitOptions.RemoveEmptyEntries).All(prefix => newNs.ContainsKey(prefix));
        var oldText = Encoding.UTF8.GetString(Read(beforeZip.GetEntry(part)!));
        var newText = Encoding.UTF8.GetString(newBytes);
        var xPrefix = !oldText.Contains("<x:worksheet", StringComparison.Ordinal) || newText.Contains("<x:worksheet", StringComparison.Ordinal);
        return $"other_bytes={(mismatches.Count == 0 ? "same" : string.Join(',', mismatches))} content_types={(mismatches.Contains("[Content_Types].xml") ? "changed" : "same")} other_xml={(preserved ? "same" : "CHANGED")} root_ns={(nsPreserved ? "same" : "CHANGED")} mc={(mcValid ? "valid" : "INVALID")} x_prefix={(xPrefix ? "same" : "CHANGED")} cell={(observed ? "ok" : "LOST")}";
    }

    private static string CellSignature(XElement cell) => ElementSignature(cell, null);

    private static string ElementSignature(XElement element, string? omittedCell)
    {
        if (element.Name == Main + "c" && (string?)element.Attribute("r") == omittedCell) return "";
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
            .OrderBy(attribute => attribute.Name.ToString()).Select(attribute => $"{attribute.Name}={attribute.Value}");
        var children = element.Nodes().Select(node => node is XElement child ? ElementSignature(child, omittedCell) : node.ToString());
        return $"<{element.Name} {string.Join(';', attributes)}>{string.Concat(children)}</{element.Name}>";
    }

    private static byte[] Read(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static MemoryStream NormalizeContentTypesForSdk(string source)
    {
        var memory = new MemoryStream();
        using (var input = ZipFile.OpenRead(source))
        using (var destination = new ZipArchive(memory, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                var target = destination.CreateEntry(entry.FullName);
                using var targetStream = target.Open();
                if (entry.FullName == "[Content_Types].xml")
                {
                    var root = XDocument.Parse(Encoding.UTF8.GetString(Read(entry))).Root!;
                    var normalized = new XElement(root.Name,
                        new XAttribute("xmlns", root.Name.NamespaceName),
                        root.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration), root.Nodes());
                    using var writer = new StreamWriter(targetStream, new UTF8Encoding(false), leaveOpen: true);
                    writer.Write(normalized.ToString(SaveOptions.DisableFormatting));
                }
                else
                {
                    using var stream = entry.Open();
                    stream.CopyTo(targetStream);
                }
            }
        }
        memory.Position = 0;
        return memory;
    }

    private static void CreateSynthetic(string source, string output, bool invalid)
    {
        using var input = ZipFile.OpenRead(source);
        using var destination = new ZipArchive(File.Create(output), ZipArchiveMode.Create);
        foreach (var entry in input.Entries)
        {
            var target = destination.CreateEntry(entry.FullName);
            using var targetStream = target.Open();
            if (entry.FullName != "xl/worksheets/sheet1.xml")
            {
                using var stream = entry.Open();
                stream.CopyTo(targetStream);
                continue;
            }
            var document = XDocument.Parse(Encoding.UTF8.GetString(Read(entry)));
            var oldRoot = document.Root!;
            var root = new XElement(oldRoot.Name,
                new XAttribute(XNamespace.Xmlns + "x", Main.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "mc", Mc.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "x14ac", X14ac.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "z", Unknown.NamespaceName),
                new XAttribute(Mc + "Ignorable", invalid ? "x14ac z missing" : "x14ac z"),
                oldRoot.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration),
                oldRoot.Nodes());
            root.Element(Main + "sheetFormatPr")!.SetAttributeValue(X14ac + "dyDescent", "0.25");
            root.Add(new XElement(Main + "extLst", new XElement(Main + "ext", new XAttribute("uri", "urn:docloupe:spike"), new XElement(Unknown + "opaque", "keep"))));
            using var writer = new StreamWriter(targetStream, new UTF8Encoding(false), leaveOpen: true);
            writer.Write(root.ToString(SaveOptions.DisableFormatting));
        }
    }
}
