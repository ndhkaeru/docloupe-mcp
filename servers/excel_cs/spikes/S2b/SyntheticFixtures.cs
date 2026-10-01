using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

internal static class SyntheticFixtures
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Markup = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace X14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private static readonly XNamespace Unknown = "urn:docloupe:spike:unknown";

    public static void NormalizeContentTypes(string sourceRoot, string outputRoot)
    {
        Directory.CreateDirectory(outputRoot);
        foreach (var file in new[] { "02-table-metadata-source.xlsx", "05-advanced-package-source.xlsm", "07-real-package-source.xlsx" })
        {
            using var input = ZipFile.OpenRead(Path.Combine(sourceRoot, file));
            using var destination = new ZipArchive(File.Create(Path.Combine(outputRoot, file)), ZipArchiveMode.Create);
            foreach (var entry in input.Entries)
            {
                var target = destination.CreateEntry(entry.FullName);
                using var targetStream = target.Open();
                if (entry.FullName != "[Content_Types].xml")
                {
                    using var sourceStream = entry.Open();
                    sourceStream.CopyTo(targetStream);
                    continue;
                }
                using var reader = new StreamReader(entry.Open());
                var root = XDocument.Parse(reader.ReadToEnd()).Root!;
                var normalized = new XElement(root.Name, new XAttribute("xmlns", root.Name.NamespaceName),
                    root.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration), root.Nodes());
                using var writer = new StreamWriter(targetStream, new UTF8Encoding(false), leaveOpen: true);
                writer.Write(normalized.ToString(SaveOptions.DisableFormatting));
            }
        }
    }

    public static void CorruptX14acNamespace(string source, string output)
    {
        using var input = ZipFile.OpenRead(source);
        using var destination = new ZipArchive(File.Create(output), ZipArchiveMode.Create);
        foreach (var entry in input.Entries)
        {
            var target = destination.CreateEntry(entry.FullName);
            using var targetStream = target.Open();
            if (entry.FullName != "xl/worksheets/sheet1.xml")
            {
                using var sourceStream = entry.Open();
                sourceStream.CopyTo(targetStream);
                continue;
            }
            using var reader = new StreamReader(entry.Open());
            var before = reader.ReadToEnd();
            var after = before.Replace("x14ac:dyDescent", "dyDescent", StringComparison.Ordinal);
            if (before == after) throw new InvalidOperationException("EX-03 attribute not present");
            using var writer = new StreamWriter(targetStream, new UTF8Encoding(false), leaveOpen: true);
            writer.Write(after);
        }
    }

    public static void Create(string source, string output, bool invalid)
    {
        using var input = ZipFile.OpenRead(source);
        using var destination = new ZipArchive(File.Create(output), ZipArchiveMode.Create);
        foreach (var entry in input.Entries)
        {
            var target = destination.CreateEntry(entry.FullName);
            using var targetStream = target.Open();
            if (entry.FullName != "xl/worksheets/sheet1.xml")
            {
                using var sourceStream = entry.Open();
                sourceStream.CopyTo(targetStream);
                continue;
            }
            using var reader = new StreamReader(entry.Open());
            var root = XDocument.Parse(reader.ReadToEnd()).Root!;
            var rebuilt = new XElement(root.Name,
                new XAttribute(XNamespace.Xmlns + "x", Main.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "mc", Markup.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "x14ac", X14ac.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "z", Unknown.NamespaceName),
                new XAttribute(Markup + "Ignorable", invalid ? "x14ac z missing" : "x14ac z"),
                root.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration), root.Nodes());
            rebuilt.Element(Main + "sheetFormatPr")!.SetAttributeValue(X14ac + "dyDescent", "0.25");
            rebuilt.Add(new XElement(Main + "extLst", new XElement(Main + "ext", new XAttribute("uri", "urn:docloupe:spike"),
                new XElement(Unknown + "opaque", "keep"))));
            using var writer = new StreamWriter(targetStream, new UTF8Encoding(false), leaveOpen: true);
            writer.Write("<?xml version='1.0' encoding='utf-8' standalone='yes'?>");
            writer.Write(rebuilt.ToString(SaveOptions.DisableFormatting));
        }
    }
}
