using System.IO.Compression;
using System.Text;

if (args is not [var output]) throw new ArgumentException("Usage: <output-directory>");
foreach (var path in SyntheticFixtures.Create(output)) Console.WriteLine(path);

public static class SyntheticFixtures
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Package = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string Types = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

    public static IReadOnlyList<string> Create(string directory)
    {
        Directory.CreateDirectory(directory);
        var variants = new[]
        {
            (Name: "default", Prefix: "", TypesPrefix: "", Bom: false, CrLf: false, Alternate: false, NoShared: false, Nested: false),
            (Name: "prefixed-x", Prefix: "x:", TypesPrefix: "ns0:", Bom: false, CrLf: false, Alternate: false, NoShared: false, Nested: false),
            (Name: "bom-crlf-standalone", Prefix: "", TypesPrefix: "ns0:", Bom: true, CrLf: true, Alternate: false, NoShared: false, Nested: false),
            (Name: "opc-percent-case", Prefix: "x:", TypesPrefix: "", Bom: false, CrLf: true, Alternate: true, NoShared: false, Nested: false),
            (Name: "new-shared-strings", Prefix: "", TypesPrefix: "", Bom: false, CrLf: false, Alternate: false, NoShared: true, Nested: false),
            (Name: "nested-workbook", Prefix: "", TypesPrefix: "", Bom: false, CrLf: false, Alternate: false, NoShared: true, Nested: true)
        };
        var paths = new List<string>();
        foreach (var variant in variants)
        {
            var path = Path.Combine(directory, variant.Name + ".xlsx");
            var prefix = variant.Prefix;
            var workbookName = variant.Alternate ? "XL/Workbook.XML" : variant.Nested ? "xl/nested/workbook.xml" : "xl/workbook.xml";
            var sheetName = variant.Alternate ? "XL/worksheets/sheet1.xml" : "xl/worksheets/sheet1.xml";
            var sstName = variant.Alternate ? "XL/sharedStrings.xml" : "xl/sharedStrings.xml";
            var workbookRels = variant.Alternate ? "XL/_rels/Workbook.XML.rels" : variant.Nested ? "xl/nested/_rels/workbook.xml.rels" : "xl/_rels/workbook.xml.rels";
            var declaration = variant.Bom || variant.CrLf ? "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" : "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
            var ns = prefix.Length == 0 ? $"xmlns=\"{Main}\"" : $"xmlns:x=\"{Main}\"";
            var book = declaration + $"<{prefix}workbook {ns} xmlns:r=\"{Office}\"><{prefix}sheets><{prefix}sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></{prefix}sheets><{prefix}calcPr calcId=\"191029\"/></{prefix}workbook>";
            var sheet = declaration + $"<{prefix}worksheet {ns} xmlns:mc=\"{Mc}\" xmlns:x14ac=\"{X14}\" mc:Ignorable=\"x14ac\"><{prefix}dimension ref=\"A1:D3\"/><{prefix}sheetViews/><{prefix}sheetFormatPr x14ac:dyDescent=\"0.25\"/><{prefix}sheetData><{prefix}row r=\"1\" x14ac:dyDescent=\"0.25\"><{prefix}c r=\"A1\" t=\"s\"><{prefix}v>0</{prefix}v></{prefix}c><{prefix}c r=\"B1\" t=\"n\"><{prefix}v>42</{prefix}v></{prefix}c><{prefix}c r=\"C1\"><{prefix}f>1+1</{prefix}f><{prefix}v>2</{prefix}v></{prefix}c></{prefix}row><{prefix}row r=\"3\"><{prefix}c r=\"D3\" t=\"inlineStr\"><{prefix}is><{prefix}t>old</{prefix}t></{prefix}is></{prefix}c></{prefix}row></{prefix}sheetData></{prefix}worksheet>";
            if (variant.NoShared) sheet = sheet.Replace($"<{prefix}c r=\"A1\" t=\"s\"><{prefix}v>0</{prefix}v></{prefix}c>", $"<{prefix}c r=\"A1\" t=\"inlineStr\"><{prefix}is><{prefix}t>hello</{prefix}t></{prefix}is></{prefix}c>", StringComparison.Ordinal);
            var shared = declaration + $"<{prefix}sst {ns} count=\"1\" uniqueCount=\"1\"><{prefix}si><{prefix}t>hello</{prefix}t><{prefix}rPh sb=\"0\" eb=\"5\"><{prefix}t>he</{prefix}t></{prefix}rPh><{prefix}phoneticPr fontId=\"0\"/></{prefix}si></{prefix}sst>";
            var typesPrefix = variant.TypesPrefix;
            var typesNs = typesPrefix.Length == 0 ? $"xmlns=\"{Types}\"" : $"xmlns:ns0=\"{Types}\"";
            var contentTypes = declaration + $"<{typesPrefix}Types {typesNs}><{typesPrefix}Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><{typesPrefix}Default Extension=\"xml\" ContentType=\"application/xml\"/><{typesPrefix}Override PartName=\"/{workbookName}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><{typesPrefix}Override PartName=\"/{sheetName}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><{typesPrefix}Override PartName=\"/{sstName}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/></{typesPrefix}Types>";
            if (variant.NoShared) contentTypes = contentTypes.Replace($"<{typesPrefix}Override PartName=\"/{sstName}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>", "", StringComparison.Ordinal);
            var mainTarget = variant.Alternate ? "/xl/WORKBOOK.XML" : "/" + workbookName;
            var rootRelationships = declaration + $"<Relationships xmlns=\"{Package}\"><Relationship Id=\"rId1\" Type=\"{Office}/officeDocument\" Target=\"{mainTarget}\"/></Relationships>";
            var worksheetTarget = variant.Alternate ? "worksheets/sheet%31.xml" : variant.Nested ? "/xl/worksheets/sheet1.xml" : "worksheets/sheet1.xml";
            var workbookRelationships = declaration + $"<Relationships xmlns=\"{Package}\"><Relationship Id=\"rId1\" Type=\"{Office}/worksheet\" Target=\"{worksheetTarget}\"/><Relationship Id=\"rId2\" Type=\"{Office}/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>";
            if (variant.NoShared) workbookRelationships = workbookRelationships.Replace($"<Relationship Id=\"rId2\" Type=\"{Office}/sharedStrings\" Target=\"sharedStrings.xml\"/>", "", StringComparison.Ordinal);
            var newline = variant.CrLf ? "\r\n" : "\n";
            using (var archive = new ZipArchive(File.Create(path), ZipArchiveMode.Create))
            {
                Write(archive, "[Content_Types].xml", contentTypes, variant.Bom, newline);
                Write(archive, "_rels/.rels", rootRelationships, false, newline);
                Write(archive, workbookName, book, variant.Bom, newline);
                Write(archive, workbookRels, workbookRelationships, false, newline);
                Write(archive, sheetName, sheet, variant.Bom, newline);
                if (!variant.NoShared) Write(archive, sstName, shared, variant.Bom, newline);
            }
            paths.Add(path);
        }
        return paths;
    }

    private static void Write(ZipArchive archive, string part, string xml, bool bom, string newline)
    {
        var entry = archive.CreateEntry(part, CompressionLevel.Optimal);
        entry.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
        using var stream = entry.Open();
        var bytes = new UTF8Encoding(bom).GetPreamble().Concat(Encoding.UTF8.GetBytes(xml.Replace("><", ">" + newline + "<", StringComparison.Ordinal))).ToArray();
        stream.Write(bytes);
    }
}
