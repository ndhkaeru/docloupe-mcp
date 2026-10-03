using System.IO.Compression;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace DocLoupe.Excel.Schema;

public sealed record SchemaIssue(string Part, string Code, string Detail);
public sealed record SchemaReport(IReadOnlyList<SchemaIssue> Issues, IReadOnlyList<SchemaIssue> Gaps);

public static class DetachedValidator
{
    public static SchemaReport CheckPackage(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var issues = new List<SchemaIssue>();
        var gaps = new List<SchemaIssue>();
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        foreach (var part in archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && !entry.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)))
        {
            try
            {
                if (part.FullName.Equals("docProps/core.xml", StringComparison.OrdinalIgnoreCase))
                {
                    var core = CheckCoreProperties(part);
                    issues.AddRange(core.Issues);
                    gaps.AddRange(core.Gaps);
                    continue;
                }
                foreach (var error in validator.Validate(Detach(part)))
                    issues.Add(new SchemaIssue(part.FullName, "SCHEMA_ERROR", Key(error)));
            }
            catch (NotSupportedException exception)
            {
                gaps.Add(new SchemaIssue(part.FullName, "G2_UNSUPPORTED_ROOT", exception.Message));
            }
            catch (XmlException exception)
            {
                issues.Add(new SchemaIssue(part.FullName, "INVALID_XML", exception.Message));
            }
        }
        return new SchemaReport(issues, gaps);
    }

    public static SchemaReport ComparePackages(string source, string written)
    {
        using var before = ZipFile.OpenRead(source);
        using var after = ZipFile.OpenRead(written);
        var parts = before.Entries.Concat(after.Entries)
            .Select(entry => entry.FullName)
            .Where(part => part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                && !part.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        return Check(source, written, parts);
    }

    public static SchemaReport Check(string source, string written, IEnumerable<string> touched)
    {
        using var before = ZipFile.OpenRead(source);
        using var after = ZipFile.OpenRead(written);
        var issues = new List<SchemaIssue>();
        var gaps = new List<SchemaIssue>();
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        foreach (var part in touched.Distinct(StringComparer.OrdinalIgnoreCase).Where(part => part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && part != "[Content_Types].xml"))
        {
            var oldEntry = before.Entries.SingleOrDefault(entry => entry.FullName.Equals(part, StringComparison.OrdinalIgnoreCase));
            var newEntry = after.Entries.SingleOrDefault(entry => entry.FullName.Equals(part, StringComparison.OrdinalIgnoreCase));
            if (newEntry is null) continue;
            try
            {
                var baseline = oldEntry is null ? [] : validator.Validate(Detach(oldEntry)).Select(Key).ToArray();
                var result = validator.Validate(Detach(newEntry)).Select(Key).ToArray();
                var remaining = baseline.GroupBy(key => key).ToDictionary(group => group.Key, group => group.Count());
                foreach (var error in result)
                {
                    if (remaining.TryGetValue(error, out var count) && count > 0) remaining[error] = count - 1;
                    else issues.Add(new SchemaIssue(part, "NEW_SCHEMA_ERROR", error));
                }
                if (oldEntry is not null)
                {
                    var originalParents = ChildLists(oldEntry);
                    var writtenParents = ChildLists(newEntry);
                    foreach (var baselineError in baseline)
                    {
                        var location = baselineError.Split('|')[0];
                        if (originalParents.TryGetValue(location, out var oldChildren)
                            && (!writtenParents.TryGetValue(location, out var newChildren) || !oldChildren.SequenceEqual(newChildren)))
                            gaps.Add(new SchemaIssue(part, "G2_MASKED_BY_BASELINE_ERROR", location));
                    }
                }
            }
            catch (NotSupportedException exception)
            {
                gaps.Add(new SchemaIssue(part, "G2_UNSUPPORTED_ROOT", exception.Message));
            }
        }
        return new SchemaReport(issues, gaps);
    }

    private static SchemaReport CheckCoreProperties(ZipArchiveEntry part)
    {
        const string coreNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        const string dublinNamespace = "http://purl.org/dc/elements/1.1/";
        var issues = new List<SchemaIssue>();
        var gaps = new List<SchemaIssue>();
        var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
        using (var stream = part.Open())
        using (var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null
        }))
            document.Load(reader);
        var root = document.DocumentElement;
        if (root is null || root.LocalName != "coreProperties" || root.NamespaceURI != coreNamespace)
            return new SchemaReport([new SchemaIssue(part.FullName, "CORE_INVALID_ROOT", "Invalid core properties root")], []);
        if (root.Attributes.OfType<XmlAttribute>().Any(attribute => attribute.NamespaceURI != "http://www.w3.org/2000/xmlns/"))
            gaps.Add(new SchemaIssue(part.FullName, "G2_CORE_UNSUPPORTED_ATTRIBUTE", "Core properties root attributes"));
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (XmlNode node in root.ChildNodes)
        {
            if (node is not XmlElement element)
            {
                if (node is XmlText text && !string.IsNullOrWhiteSpace(text.Value))
                    issues.Add(new SchemaIssue(part.FullName, "CORE_INVALID_TEXT", "Text outside core property"));
                continue;
            }
            var dublin = element.LocalName is "title" or "subject" or "creator" or "description";
            var supported = dublin || element.LocalName is "keywords" or "category" or "contentStatus" or "lastModifiedBy";
            if (!supported || element.NamespaceURI != (dublin ? dublinNamespace : coreNamespace))
            {
                gaps.Add(new SchemaIssue(part.FullName, "G2_CORE_UNSUPPORTED_FIELD", element.LocalName));
                continue;
            }
            if (!seen.Add(element.LocalName))
                issues.Add(new SchemaIssue(part.FullName, "CORE_DUPLICATE_FIELD", element.LocalName));
            if (element.Attributes.OfType<XmlAttribute>().Any(attribute => attribute.NamespaceURI != "http://www.w3.org/2000/xmlns/"))
                gaps.Add(new SchemaIssue(part.FullName, "G2_CORE_UNSUPPORTED_ATTRIBUTE", element.LocalName));
            if (element.ChildNodes.OfType<XmlElement>().Any() || element.InnerText.Length > 4096)
                issues.Add(new SchemaIssue(part.FullName, "CORE_INVALID_VALUE", element.LocalName));
        }
        return new SchemaReport(issues, gaps);
    }

    private static string Key(ValidationErrorInfo error) => $"{error.Path?.XPath}|{error.Id}|{error.Description}";

    private static Dictionary<string, string[]> ChildLists(ZipArchiveEntry part)
    {
        var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
        using (var stream = part.Open())
        using (var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
            document.Load(reader);
        var lists = new Dictionary<string, string[]>(StringComparer.Ordinal);
        void Walk(XmlElement element, string path)
        {
            var children = element.ChildNodes.OfType<XmlElement>().ToArray();
            lists[path] = children.Select(child => child.NamespaceURI + ":" + child.LocalName).ToArray();
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var child in children)
            {
                var name = child.NamespaceURI + ":" + child.LocalName;
                counts[name] = counts.GetValueOrDefault(name) + 1;
                Walk(child, path + "/x:" + child.LocalName + "[" + counts[name] + "]");
            }
        }
        var root = document.DocumentElement ?? throw new InvalidDataException("Missing XML root");
        Walk(root, "/x:" + root.LocalName + "[1]");
        return lists;
    }

    private static OpenXmlElement Detach(ZipArchiveEntry part)
    {
        var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
        using (var stream = part.Open())
        using (var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
            document.Load(reader);
        var root = document.DocumentElement ?? throw new InvalidDataException("Missing XML root");
        OpenXmlElement detached = (root.LocalName, root.NamespaceURI) switch
        {
            ("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") => new Worksheet(),
            ("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") => new Workbook(),
            ("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") => new SharedStringTable(),
            _ => throw new NotSupportedException($"Detached validation not implemented for {part.FullName}: {{{root.NamespaceURI}}}{root.LocalName}")
        };
        foreach (XmlAttribute attribute in root.Attributes)
        {
            if (attribute.Prefix == "xmlns") detached.AddNamespaceDeclaration(attribute.LocalName, attribute.Value);
            else if (attribute.Name != "xmlns")
                detached.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceURI, attribute.Value));
        }
        detached.InnerXml = root.InnerXml;
        return detached;
    }
}
