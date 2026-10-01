using DocLoupe.Excel.Engine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

foreach (var (name, factory) in new (string Name, Func<OpenXmlCompositeElement> Factory)[]
{
    ("worksheet", () => new Worksheet()),
    ("workbook", () => new Workbook()),
    ("sheetData", () => new SheetData()),
    ("row", () => new Row()),
    ("cell", () => new Cell())
})
{
    var candidates = typeof(Worksheet).Assembly.GetTypes()
        .Where(type => type.Namespace == typeof(Worksheet).Namespace && !type.IsAbstract && typeof(OpenXmlElement).IsAssignableFrom(type)
            && type.GetConstructor(Type.EmptyTypes) is not null)
        .OrderBy(type => type.FullName, StringComparer.Ordinal)
        .Select(type => (OpenXmlElement)Activator.CreateInstance(type)!)
        .Where(candidate => factory().AddChild(candidate.CloneNode(true), false))
        .DistinctBy(candidate => (candidate.NamespaceUri, candidate.LocalName))
        .ToArray();
    var root = factory();
    foreach (var candidate in candidates.Reverse())
        root.AddChild(candidate, true);
    var generated = root.ChildElements.Select(element => element.LocalName).ToArray();
    Console.WriteLine(name + "=" + string.Join(',', generated));
    var expected = name switch
    {
        "worksheet" => SchemaElementOrder.Worksheet, "workbook" => SchemaElementOrder.Workbook,
        "sheetData" => SchemaElementOrder.SheetData, "row" => SchemaElementOrder.Row,
        "cell" => SchemaElementOrder.Cell, _ => throw new InvalidOperationException("Unknown root")
    };
    if (!expected.SequenceEqual(generated))
        throw new InvalidDataException($"The {name} element order no longer matches the SDK schema metadata");
}
