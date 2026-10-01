using System.Xml;

namespace DocLoupe.Excel.Verify;

public sealed record MarkupIssue(string Code, string Detail);

public sealed record MarkupResult(IReadOnlyList<MarkupIssue> Issues)
{
    public bool Passed => Issues.Count == 0;
}

public static class MarkupCompatibilityVerifier
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string MarkupNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    public static MarkupResult Check(Stream written, Stream? original = null, string? editedCell = null)
    {
        var issues = new List<MarkupIssue>();
        try
        {
            var after = Inspect(written, issues);
            if (original is not null)
            {
                var before = Inspect(original, new List<MarkupIssue>());
                Compare(before, after, editedCell, issues);
            }
        }
        catch (XmlException exception)
        {
            issues.Add(new MarkupIssue("INVALID_XML", exception.Message));
        }

        return new MarkupResult(issues);
    }

    private static Snapshot Inspect(Stream stream, List<MarkupIssue> issues)
    {
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreWhitespace = false
        });
        var shapes = new List<Shape>();
        var rootNamespaces = new Dictionary<string, string>(StringComparer.Ordinal);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element)
                continue;

            var attributes = new List<AttributeShape>();
            var isCell = reader.LocalName == "c" && reader.NamespaceURI == MainNamespace;
            var cellReference = isCell ? reader.GetAttribute("r") : null;
            var isChoice = reader.LocalName == "Choice" && reader.NamespaceURI == MarkupNamespace;
            if (reader.HasAttributes)
            {
                reader.MoveToFirstAttribute();
                do
                {
                    var declaration = reader.Prefix == "xmlns" || reader.Name == "xmlns";
                    if (shapes.Count == 0 && declaration)
                        rootNamespaces[reader.Prefix == "xmlns" ? reader.LocalName : ""] = reader.Value;
                    // mc:Choice/@Requires is an unqualified attribute whose tokens are prefixes, like mc:Ignorable.
                    var requires = isChoice && reader.LocalName == "Requires" && reader.NamespaceURI.Length == 0;
                    if (requires || reader.NamespaceURI == MarkupNamespace && reader.LocalName is "Ignorable" or "ProcessContent" or "MustUnderstand")
                        CheckPrefixes(reader, issues);
                    attributes.Add(new AttributeShape(reader.Prefix, reader.LocalName, reader.NamespaceURI,
                        declaration || requires || reader.NamespaceURI == MarkupNamespace ? reader.Value : null));
                }
                while (reader.MoveToNextAttribute());
                reader.MoveToElement();
            }

            shapes.Add(new Shape(reader.Depth, reader.Prefix, reader.LocalName, reader.NamespaceURI, cellReference,
                attributes.OrderBy(attribute => attribute.Namespace).ThenBy(attribute => attribute.Local).ToArray()));
        }

        if (shapes.Count == 0)
            issues.Add(new MarkupIssue("EMPTY_XML", "No document element"));

        return new Snapshot(rootNamespaces, shapes);
    }

    private static void CheckPrefixes(XmlReader reader, List<MarkupIssue> issues)
    {
        foreach (var token in reader.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            var prefix = token.Split(':', 2)[0];
            if (prefix.Length == 0 || reader.LookupNamespace(prefix) is null)
                issues.Add(new MarkupIssue("UNDECLARED_MC_PREFIX", $"{reader.LocalName}: {token}"));
        }
    }

    private static void Compare(Snapshot before, Snapshot after, string? editedCell, List<MarkupIssue> issues)
    {
        foreach (var pair in before.RootNamespaces)
        {
            if (!after.RootNamespaces.TryGetValue(pair.Key, out var value) || value != pair.Value)
                issues.Add(new MarkupIssue("ROOT_NAMESPACE_CHANGED", pair.Key));
        }

        var oldShapes = WithoutEditedCellChildren(before.Shapes, editedCell);
        var newShapes = WithoutEditedCellChildren(after.Shapes, editedCell);
        if (oldShapes.Count != newShapes.Count)
        {
            issues.Add(new MarkupIssue("ELEMENT_SHAPE_CHANGED", $"{oldShapes.Count} -> {newShapes.Count}"));
            return;
        }

        for (var index = 0; index < oldShapes.Count; index++)
        {
            var oldShape = oldShapes[index];
            var newShape = newShapes[index];
            if (oldShape.Prefix != newShape.Prefix || oldShape.Local != newShape.Local || oldShape.Namespace != newShape.Namespace)
                issues.Add(new MarkupIssue("ELEMENT_NAMESPACE_CHANGED", $"element {index}: {oldShape.Prefix}:{oldShape.Local} -> {newShape.Prefix}:{newShape.Local}"));

            var oldAttributes = FilterAttributes(oldShape, editedCell);
            var newAttributes = FilterAttributes(newShape, editedCell);
            if (!oldAttributes.SequenceEqual(newAttributes))
                issues.Add(new MarkupIssue("ATTRIBUTE_NAMESPACE_CHANGED", $"element {index}: {oldShape.Local}"));
        }
    }

    private static IReadOnlyList<Shape> WithoutEditedCellChildren(IReadOnlyList<Shape> shapes, string? editedCell)
    {
        if (editedCell is null) return shapes;
        var skippedDepth = -1;
        var result = new List<Shape>();
        foreach (var shape in shapes)
        {
            if (skippedDepth >= 0 && shape.Depth > skippedDepth)
                continue;
            skippedDepth = -1;
            if (shape.Local == "c" && shape.Namespace == MainNamespace && shape.CellReference == editedCell)
            {
                skippedDepth = shape.Depth;
                result.Add(shape);
                continue;
            }
            result.Add(shape);
        }
        return result;
    }

    private static IEnumerable<AttributeShape> FilterAttributes(Shape shape, string? editedCell) =>
        shape.Attributes.Where(attribute => editedCell is null || shape.CellReference != editedCell || attribute.Local != "t" || attribute.Namespace.Length != 0);

    private sealed record Snapshot(Dictionary<string, string> RootNamespaces, IReadOnlyList<Shape> Shapes);
    private sealed record Shape(int Depth, string Prefix, string Local, string Namespace, string? CellReference, AttributeShape[] Attributes);
    private sealed record AttributeShape(string Prefix, string Local, string Namespace, string? Value);
}
