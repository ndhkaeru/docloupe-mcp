using System.IO.Compression;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

if (args is not [var sourceRoot, var editedRoot])
    throw new ArgumentException("Usage: <fixture-sources> <S2b-output-dir>");

var sources = Directory.GetFiles(sourceRoot)
    .Where(path => Path.GetFileName(path).Length > 2 && Path.GetFileName(path)[..2] is "00" or "01" or "02" or "03" or "04" or "05" or "06" or "07")
    .Where(path => !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal))
    .OrderBy(path => path).ToArray();
if (sources.Length != 8) throw new InvalidOperationException($"Expected eight sources, found {sources.Length}");

var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var unexpected = new List<string>();
foreach (var source in sources)
{
    var edited = Path.Combine(editedRoot, Path.GetFileNameWithoutExtension(source) + "-s2b" + Path.GetExtension(source));
    var before = Inspect(source);
    var after = Inspect(edited);
    var baselineErrors = validator.Validate(before).ToList();
    var baseline = baselineErrors.Select(error => error.Id).OrderBy(id => id).ToArray();
    var written = validator.Validate(after).Select(error => error.Id).OrderBy(id => id).ToArray();
    Console.WriteLine($"{Path.GetFileName(source)} | detached={before.Descendants<Cell>().Count()}/{after.Descendants<Cell>().Count()} baseline_errors={baseline.Length} written_errors={written.Length} same_error_ids={baseline.SequenceEqual(written)}");

    // Inject one schema error at three depths, into fresh copies of the written root, and check each one is
    // reported. The validator reports only the first content-model error per parent element, so an error
    // whose parent already has a baseline content-model error is expected to be MASKED; deeper errors are not.
    var maskedParents = baselineErrors.Select(error => error.Path?.XPath).Where(path => path is not null).ToHashSet();
    var probes = new (string Name, string ParentPath, Action<Worksheet> Inject)[]
    {
        ("worksheet-level", "/x:worksheet[1]", worksheet => worksheet.AppendChild(new Cell())),
        ("sheetData-level", "/x:worksheet[1]/x:sheetData[1]", worksheet => worksheet.GetFirstChild<SheetData>()!.AppendChild(new Cell())),
        ("inside-cell", "", worksheet => worksheet.Descendants<Cell>().First().AppendChild(new CellFormula("1+1"))),
    };
    foreach (var (name, parentPath, inject) in probes)
    {
        var probe = (Worksheet)after.CloneNode(true);
        inject(probe);
        var detected = validator.Validate(probe).Count() > written.Length;
        var expectMasked = parentPath.Length > 0 && maskedParents.Contains(parentPath);
        var outcome = detected ? "detected" : "MASKED";
        Console.WriteLine($"  injected {name,-15} | {outcome}{(expectMasked ? " (parent has a baseline content-model error)" : "")}");
        if (detected == expectMasked) unexpected.Add($"{Path.GetFileName(source)} {name}: {outcome}");
    }
}
if (unexpected.Count > 0)
    throw new InvalidOperationException("Detection did not match the masking rule: " + string.Join("; ", unexpected));

static Worksheet Inspect(string path)
{
    using var archive = ZipFile.OpenRead(path);
    var entry = archive.Entries.First(part => part.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal)
        && part.FullName.EndsWith(".xml", StringComparison.Ordinal));
    var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
    using var stream = entry.Open();
    using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
    document.Load(reader);
    var worksheet = new Worksheet();
    foreach (XmlAttribute attribute in document.DocumentElement!.Attributes)
    {
        if (attribute.Prefix == "xmlns")
            worksheet.AddNamespaceDeclaration(attribute.LocalName, attribute.Value);
        else if (attribute.Name != "xmlns")
            worksheet.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceURI, attribute.Value));
    }
    worksheet.InnerXml = document.DocumentElement.InnerXml;
    return worksheet;
}
