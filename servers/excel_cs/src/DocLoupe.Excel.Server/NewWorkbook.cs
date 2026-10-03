using System.IO.Compression;
using System.Text;
using System.Xml;

namespace DocLoupe.Excel.Server;

internal static class NewWorkbook
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string Types = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string Core = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private const string DublinCore = "http://purl.org/dc/elements/1.1/";

    public static string[] Write(string path, string[]? requestedSheets, string? activeSheet, string format,
        IReadOnlyDictionary<string, string>? coreProperties = null)
    {
        var workbookContentType = format switch
        {
            "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "xltx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "xltm" => "application/vnd.ms-excel.template.macroEnabled.main+xml",
            _ => throw new NotSupportedException("Unsupported workbook format")
        };
        var sheets = requestedSheets ?? ["Sheet1"];
        if (sheets.Length is < 1 or > 256 || sheets.Any(sheet => string.IsNullOrWhiteSpace(sheet) ||
            sheet.Length > 31 || sheet.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0 ||
            sheet.StartsWith('\'') || sheet.EndsWith('\'')))
            throw new ArgumentException("Invalid worksheet names");
        if (sheets.Distinct(StringComparer.OrdinalIgnoreCase).Count() != sheets.Length)
            throw new ArgumentException("Worksheet names must be unique ignoring case");
        foreach (var sheet in sheets) XmlConvert.VerifyXmlChars(sheet);
        var active = activeSheet is null ? 0 : Array.IndexOf(sheets, activeSheet);
        if (active < 0) throw new ArgumentException("Active worksheet does not exist");

        using var archive = new ZipArchive(File.Create(path), ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", writer =>
        {
            writer.WriteStartElement("Types", Types);
            WriteContentType(writer, "Default", "Extension", "rels", "application/vnd.openxmlformats-package.relationships+xml");
            WriteContentType(writer, "Default", "Extension", "xml", "application/xml");
            WriteContentType(writer, "Override", "PartName", "/xl/workbook.xml", workbookContentType);
            if (coreProperties is { Count: > 0 })
                WriteContentType(writer, "Override", "PartName", "/docProps/core.xml",
                    "application/vnd.openxmlformats-package.core-properties+xml");
            for (var index = 0; index < sheets.Length; index++)
                WriteContentType(writer, "Override", "PartName", $"/xl/worksheets/sheet{index + 1}.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            writer.WriteEndElement();
        });
        Add(archive, "_rels/.rels", writer =>
        {
            writer.WriteStartElement("Relationships", Relationships);
            WriteRelationship(writer, "rId1", Office + "/officeDocument", "xl/workbook.xml");
            if (coreProperties is { Count: > 0 })
                WriteRelationship(writer, "rId2", Relationships + "/metadata/core-properties", "docProps/core.xml");
            writer.WriteEndElement();
        });
        if (coreProperties is { Count: > 0 })
            Add(archive, "docProps/core.xml", writer =>
            {
                writer.WriteStartElement("cp", "coreProperties", Core);
                writer.WriteAttributeString("xmlns", "dc", null, DublinCore);
                foreach (var (name, value) in coreProperties)
                {
                    var dublinCore = name is "title" or "subject" or "creator" or "description";
                    writer.WriteElementString(dublinCore ? "dc" : "cp", name, dublinCore ? DublinCore : Core, value);
                }
                writer.WriteEndElement();
            });
        Add(archive, "xl/workbook.xml", writer =>
        {
            writer.WriteStartElement("workbook", Main);
            writer.WriteAttributeString("xmlns", "r", null, Office);
            writer.WriteStartElement("bookViews", Main);
            writer.WriteStartElement("workbookView", Main);
            writer.WriteAttributeString("activeTab", active.ToString(System.Globalization.CultureInfo.InvariantCulture));
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("sheets", Main);
            for (var index = 0; index < sheets.Length; index++)
            {
                writer.WriteStartElement("sheet", Main);
                writer.WriteAttributeString("name", sheets[index]);
                writer.WriteAttributeString("sheetId", (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
                writer.WriteAttributeString("r", "id", Office, $"rId{index + 1}");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            writer.WriteEndElement();
        });
        Add(archive, "xl/_rels/workbook.xml.rels", writer =>
        {
            writer.WriteStartElement("Relationships", Relationships);
            for (var index = 0; index < sheets.Length; index++)
                WriteRelationship(writer, $"rId{index + 1}", Office + "/worksheet",
                    $"worksheets/sheet{index + 1}.xml");
            writer.WriteEndElement();
        });
        for (var index = 0; index < sheets.Length; index++)
            Add(archive, $"xl/worksheets/sheet{index + 1}.xml", writer =>
            {
                writer.WriteStartElement("worksheet", Main);
                writer.WriteStartElement("sheetData", Main);
                writer.WriteFullEndElement();
                writer.WriteEndElement();
            });
        return sheets;
    }

    private static void Add(ZipArchive archive, string name, Action<XmlWriter> write)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false,
            CloseOutput = false
        });
        writer.WriteStartDocument();
        write(writer);
        writer.WriteEndDocument();
    }

    private static void WriteContentType(XmlWriter writer, string element, string attribute, string key, string contentType)
    {
        writer.WriteStartElement(element, Types);
        writer.WriteAttributeString(attribute, key);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteRelationship(XmlWriter writer, string id, string type, string target)
    {
        writer.WriteStartElement("Relationship", Relationships);
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Type", type);
        writer.WriteAttributeString("Target", target);
        writer.WriteEndElement();
    }
}
