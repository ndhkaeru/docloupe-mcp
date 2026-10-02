using System.IO.Compression;
using System.Text.Json.Serialization;
using System.Xml;

namespace DocLoupe.Excel.Verify;

public sealed record DeclaredByteSpan(string Part, int Start, int End, byte[] Before, byte[] After);
public sealed record FormulaCacheExpectation(string Type, string Value);
public sealed record CellExpectation(string Sheet, string Address, string Kind, string? Value, bool AllowMissing = false, bool RequireMissing = false, bool KeepCache = false, FormulaCacheExpectation? ExplicitCache = null);
public sealed record CellRead(string Address, string Kind, string? Value, string? Formula,
    [property: JsonIgnore] string? CacheType = null, [property: JsonIgnore] string? CacheRawValue = null);
public sealed record GateIssue(string Gate, string Code, string Detail);

public static class P2aGates
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string Types = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static IReadOnlyList<GateIssue> CheckPackage(string path, IEnumerable<string> editedParts, string? sourceExtension = null)
    {
        var issues = new List<GateIssue>();
        using var zip = ZipFile.OpenRead(path);
        var entries = zip.Entries.GroupBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase).ToArray();
        foreach (var group in entries.Where(group => group.Count() != 1)) issues.Add(new("G1", "DUPLICATE_PART", group.Key));
        var parts = entries.ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        if (!parts.ContainsKey("_rels/.rels") || !parts.ContainsKey("[Content_Types].xml"))
            return [new("G1", "MISSING_ROOT", "OPC root relationships or content types missing")];
        var contentTypes = Load(parts["[Content_Types].xml"]);
        if (contentTypes.DocumentElement?.LocalName != "Types" || contentTypes.DocumentElement.NamespaceURI != Types)
            issues.Add(new("G1", "INVALID_TYPES", "Wrong content-types root"));
        var defaults = contentTypes.GetElementsByTagName("Default", Types).OfType<XmlElement>()
            .ToDictionary(element => element.GetAttribute("Extension"), element => element.GetAttribute("ContentType"), StringComparer.OrdinalIgnoreCase);
        var overrides = contentTypes.GetElementsByTagName("Override", Types).OfType<XmlElement>()
            .ToDictionary(element => element.GetAttribute("PartName").TrimStart('/'), element => element.GetAttribute("ContentType"), StringComparer.OrdinalIgnoreCase);
        foreach (var part in parts.Keys.Where(part => part != "[Content_Types].xml"))
        {
            var extension = Path.GetExtension(part).TrimStart('.');
            if (!overrides.ContainsKey(part) && !defaults.ContainsKey(extension))
                issues.Add(new("G1", "MISSING_CONTENT_TYPE", part));
        }
        var relationshipMap = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        string? workbook = null;
        foreach (var (part, entry) in parts.Where(pair => pair.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            var source = SourcePart(part);
            if (source is null) { issues.Add(new("G1", "INVALID_RELS_PATH", part)); continue; }
            var doc = Load(entry);
            if (doc.DocumentElement?.LocalName != "Relationships" || doc.DocumentElement.NamespaceURI != Relationships)
            { issues.Add(new("G1", "INVALID_RELS", part)); continue; }
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var relationship in doc.GetElementsByTagName("Relationship", Relationships).OfType<XmlElement>())
            {
                var id = relationship.GetAttribute("Id");
                if (!ids.Add(id)) issues.Add(new("G1", "DUPLICATE_RELATIONSHIP", $"{part}: {id}"));
                if (relationship.GetAttribute("TargetMode") == "External") continue;
                try
                {
                    var target = Resolve(source, relationship.GetAttribute("Target"));
                    if (!parts.ContainsKey(target)) issues.Add(new("G1", "DANGLING_TARGET", $"{part}: {target}"));
                    if (source == "" && relationship.GetAttribute("Type").EndsWith("/officeDocument", StringComparison.Ordinal)) workbook = target;
                }
                catch (Exception exception) when (exception is ArgumentException or InvalidDataException or UriFormatException)
                { issues.Add(new("G1", "INVALID_TARGET", $"{part}: {exception.Message}")); }
            }
            relationshipMap[source] = ids;
        }
        if (workbook is null || !parts.ContainsKey(workbook) || !relationshipMap.ContainsKey(workbook))
            issues.Add(new("G1", "MISSING_WORKBOOK", workbook ?? "officeDocument relationship missing"));
        else
        {
            var extension = sourceExtension ?? Path.GetExtension(path);
            var type = overrides.TryGetValue(workbook, out var overridden) ? overridden : defaults.GetValueOrDefault(Path.GetExtension(workbook).TrimStart('.'));
            var expectedMarker = extension.ToLowerInvariant() switch
            {
                ".xlsm" => "macroEnabled.main+xml", ".xltm" => "template.macroEnabled.main+xml",
                ".xltx" => "template.main+xml", ".xlsx" => "sheet.main+xml", _ => null
            };
            if (expectedMarker is not null && (type is null || !type.EndsWith(expectedMarker, StringComparison.Ordinal)))
                issues.Add(new("G1", "WORKBOOK_CONTENT_TYPE_MISMATCH", $"{workbook}: {type}"));
        }
        foreach (var part in editedParts.Where(part => parts.ContainsKey(part) && part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            var document = Load(parts[part]);
            foreach (XmlElement element in document.GetElementsByTagName("*"))
            foreach (XmlAttribute attribute in element.Attributes)
                if (attribute.NamespaceURI == Office && attribute.LocalName is "id" or "embed" or "link" or "pict"
                    && (!relationshipMap.TryGetValue(part, out var ids) || !ids.Contains(attribute.Value)))
                    issues.Add(new("G1", "DANGLING_REFERENCE", $"{part}: {attribute.Value}"));
        }
        return issues;
    }

    public static IReadOnlyList<GateIssue> CheckPreservation(string source, string written,
        IEnumerable<DeclaredByteSpan> declared, IEnumerable<string> addedOrRemoved)
    {
        var issues = new List<GateIssue>();
        using var before = ZipFile.OpenRead(source);
        using var after = ZipFile.OpenRead(written);
        var oldParts = before.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var newParts = after.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var allowedParts = addedOrRemoved.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var byPart = declared.GroupBy(edit => edit.Part, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.OrderBy(edit => edit.Start).ThenBy(edit => edit.End).ToArray(), StringComparer.OrdinalIgnoreCase);
        foreach (var part in oldParts.Keys.Union(newParts.Keys, StringComparer.OrdinalIgnoreCase))
        {
            if (!oldParts.TryGetValue(part, out var original) || !newParts.TryGetValue(part, out var result))
            {
                if (!allowedParts.Contains(part)) issues.Add(new("G5", "UNDECLARED_PART", part));
                continue;
            }
            var sourceBytes = Read(original);
            var writtenBytes = Read(result);
            if (!byPart.TryGetValue(part, out var edits))
            {
                if (!sourceBytes.AsSpan().SequenceEqual(writtenBytes)) issues.Add(new("G5", "UNDECLARED_PART_CHANGE", part));
                continue;
            }
            var sourceOffset = 0;
            var writtenOffset = 0;
            foreach (var edit in edits)
            {
                if (edit.Start < sourceOffset || edit.End > sourceBytes.Length || edit.End < edit.Start
                    || !sourceBytes.AsSpan(edit.Start, edit.End - edit.Start).SequenceEqual(edit.Before))
                { issues.Add(new("G5", "INVALID_DECLARATION", part)); break; }
                var length = edit.Start - sourceOffset;
                if (writtenOffset + length + edit.After.Length > writtenBytes.Length
                    || !sourceBytes.AsSpan(sourceOffset, length).SequenceEqual(writtenBytes.AsSpan(writtenOffset, length))
                    || !writtenBytes.AsSpan(writtenOffset + length, edit.After.Length).SequenceEqual(edit.After))
                { issues.Add(new("G5", "UNDECLARED_BYTES", part)); break; }
                sourceOffset = edit.End;
                writtenOffset += length + edit.After.Length;
            }
            if (!sourceBytes.AsSpan(sourceOffset).SequenceEqual(writtenBytes.AsSpan(writtenOffset)))
                issues.Add(new("G5", "UNDECLARED_SUFFIX", part));
        }
        return issues;
    }

    public static IReadOnlyList<GateIssue> CheckTouchedCells(string source, string written, IEnumerable<CellExpectation> expected)
    {
        var issues = new List<GateIssue>();
        using var oldZip = ZipFile.OpenRead(source);
        using var newZip = ZipFile.OpenRead(written);
        var oldEntries = oldZip.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var newEntries = newZip.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var main = Resolve("", Relations(oldEntries, "").Values.Single(item => item.Type.EndsWith("/officeDocument", StringComparison.Ordinal)).Target);
        var workbook = Load(oldEntries[main]);
        var relationships = Relations(oldEntries, main);
        foreach (var group in expected.GroupBy(item => item.Sheet))
        {
            var sheet = workbook.GetElementsByTagName("sheet", Main).OfType<XmlElement>().Single(item => item.GetAttribute("name") == group.Key);
            var part = Resolve(main, relationships[sheet.GetAttribute("id", Office)].Target);
            var oldSheet = Load(oldEntries[part]);
            var newSheet = Load(newEntries[part]);
            foreach (var cell in group)
            {
                var before = oldSheet.GetElementsByTagName("c", Main).OfType<XmlElement>().SingleOrDefault(item => item.GetAttribute("r") == cell.Address);
                var after = newSheet.GetElementsByTagName("c", Main).OfType<XmlElement>().SingleOrDefault(item => item.GetAttribute("r") == cell.Address);
                if (after is null)
                {
                    if (before is null && cell.AllowMissing || before is not null && cell.RequireMissing) continue;
                    issues.Add(new("G5", "CELL_MISSING", $"{group.Key}!{cell.Address}"));
                    continue;
                }
                if (cell.RequireMissing)
                {
                    issues.Add(new("G5", "CELL_NOT_REMOVED", $"{group.Key}!{cell.Address}"));
                    continue;
                }
                if (cell.ExplicitCache is { } expectedCache)
                {
                    var cache = after.ChildNodes.OfType<XmlElement>()
                        .Where(child => child.NamespaceURI == Main && child.LocalName == "v").ToArray();
                    if (after.GetAttribute("t") != (expectedCache.Type == "n" ? "" : expectedCache.Type) ||
                        cache.Length != 1 || cache[0].InnerText != expectedCache.Value)
                        issues.Add(new("G5", "FORMULA_CACHE_MISMATCH", $"{group.Key}!{cell.Address}"));
                }
                if (before is null)
                {
                    if (cell.AllowMissing) issues.Add(new("G5", "UNEXPECTED_CELL_CREATED", $"{group.Key}!{cell.Address}"));
                    else if (cell.KeepCache) issues.Add(new("G5", "FORMULA_CACHE_SOURCE_MISSING", $"{group.Key}!{cell.Address}"));
                    continue;
                }
                if (cell.KeepCache)
                {
                    static string[] Cache(XmlElement element) => element.ChildNodes.OfType<XmlElement>()
                        .Where(child => child.NamespaceURI == Main && child.LocalName == "v")
                        .Select(child => child.OuterXml).ToArray();
                    if (before.GetAttribute("t") != after.GetAttribute("t") || !Cache(before).SequenceEqual(Cache(after)))
                        issues.Add(new("G5", "FORMULA_CACHE_CHANGED", $"{group.Key}!{cell.Address}"));
                }
                static string[] UnchangedAttributes(XmlElement element) => element.Attributes.OfType<XmlAttribute>()
                    .Where(attribute => attribute.NamespaceURI != "http://www.w3.org/2000/xmlns/"
                        && (attribute.NamespaceURI != "" || attribute.LocalName != "t"))
                    .Select(attribute => attribute.NamespaceURI + ":" + attribute.LocalName + "=" + attribute.Value)
                    .OrderBy(attribute => attribute, StringComparer.Ordinal).ToArray();
                if (!UnchangedAttributes(before).SequenceEqual(UnchangedAttributes(after)))
                    issues.Add(new("G5", "CELL_ATTRIBUTE_CHANGED", $"{group.Key}!{cell.Address}"));
                if (before.ChildNodes.OfType<XmlElement>().Concat(after.ChildNodes.OfType<XmlElement>())
                    .Any(child => child.NamespaceURI != Main || child.LocalName is not ("f" or "v" or "is")))
                    issues.Add(new("G5", "CELL_UNMODELED_CHILD", $"{group.Key}!{cell.Address}"));
            }
        }
        return issues;
    }

    public static IReadOnlyList<CellRead> ReadCells(string path, string sheetName, IEnumerable<string> addresses)
    {
        using var archive = ZipFile.OpenRead(path);
        var entries = archive.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var root = Relations(entries, "");
        var main = Resolve("", root.Values.Single(item => item.Type.EndsWith("/officeDocument", StringComparison.Ordinal)).Target);
        var workbook = Load(entries[main]);
        var relationships = Relations(entries, main);
        var sheet = workbook.GetElementsByTagName("sheet", Main).OfType<XmlElement>()
            .Single(item => item.GetAttribute("name") == sheetName);
        var part = Resolve(main, relationships[sheet.GetAttribute("id", Office)].Target);
        var document = Load(entries[part]);
        var sst = relationships.Values.FirstOrDefault(item => item.Type.EndsWith("/sharedStrings", StringComparison.Ordinal));
        string[] shared = sst.Type is null ? [] : Load(entries[Resolve(main, sst.Target)])
            .GetElementsByTagName("si", Main).OfType<XmlElement>()
            .Select(TextValue).ToArray();
        var requested = addresses.ToHashSet(StringComparer.Ordinal);
        return document.GetElementsByTagName("c", Main).OfType<XmlElement>()
            .Where(item => requested.Contains(item.GetAttribute("r"))).Select(item =>
            {
                var type = item.GetAttribute("t");
                var scalar = item.GetElementsByTagName("v", Main).OfType<XmlElement>().FirstOrDefault()?.InnerText;
                var formula = item.GetElementsByTagName("f", Main).OfType<XmlElement>().FirstOrDefault()?.InnerText;
                var value = type switch
                {
                    "s" when int.TryParse(scalar, out var index) && index >= 0 && index < shared.Length => shared[index],
                    "inlineStr" => item.ChildNodes.OfType<XmlElement>().FirstOrDefault(child => child.LocalName == "is" && child.NamespaceURI == Main) is { } inline ? TextValue(inline) : null,
                    "b" => scalar == "1" ? "true" : "false",
                    "e" => scalar,
                    _ => scalar
                };
                return new CellRead(item.GetAttribute("r"), formula is not null ? "formula" : type switch
                {
                    "s" => "text", "inlineStr" => "inline", "b" => "boolean", "e" => "error",
                    _ when value is null => "blank", _ => "number"
                }, value, formula, formula is null ? null : type, formula is null ? null : scalar);
            }).ToArray();
    }

    public static IReadOnlyList<GateIssue> CheckIntent(string path, IEnumerable<CellExpectation> expected)
    {
        var issues = new List<GateIssue>();
        using var zip = ZipFile.OpenRead(path);
        var entries = zip.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var rootRel = Relations(entries, "");
        var main = Resolve("", rootRel.Values.Single(relationship => relationship.Type.EndsWith("/officeDocument", StringComparison.Ordinal)).Target);
        var workbook = Load(entries[main]);
        var sheetRels = Relations(entries, main);
        var sharedRelationship = sheetRels.Values.FirstOrDefault(relationship => relationship.Type.EndsWith("/sharedStrings", StringComparison.Ordinal));
        string[] shared = sharedRelationship.Type is null ? [] : Load(entries[Resolve(main, sharedRelationship.Target)])
            .GetElementsByTagName("si", Main).OfType<XmlElement>()
            .Select(TextValue).ToArray();
        foreach (var group in expected.GroupBy(item => item.Sheet))
        {
            var sheet = workbook.GetElementsByTagName("sheet", Main).OfType<XmlElement>().Single(element => element.GetAttribute("name") == group.Key);
            var id = sheet.GetAttribute("id", Office);
            var part = Resolve(main, sheetRels[id].Target);
            var document = Load(entries[part]);
            foreach (var expectation in group)
            {
                var cell = document.GetElementsByTagName("c", Main).OfType<XmlElement>()
                    .SingleOrDefault(item => item.GetAttribute("r") == expectation.Address);
                if (cell is null)
                {
                    if (expectation.AllowMissing && expectation.Kind == "blank") continue;
                    issues.Add(new("G4", "INTENT_MISSING", $"{group.Key}!{expectation.Address}"));
                    continue;
                }
                if (expectation.RequireMissing)
                {
                    issues.Add(new("G4", "INTENT_PRESENT", $"{group.Key}!{expectation.Address}: cell should have been removed"));
                    continue;
                }
                if (expectation.AllowMissing && expectation.Kind == "blank" &&
                    (cell.HasAttribute("t") || cell.ChildNodes.OfType<XmlElement>().Any()))
                {
                    issues.Add(new("G4", "INTENT_MISMATCH", $"{group.Key}!{expectation.Address}: clear left cell content or type"));
                    continue;
                }
                var formula = cell.GetElementsByTagName("f", Main).OfType<XmlElement>().FirstOrDefault();
                var cached = cell.GetElementsByTagName("v", Main).OfType<XmlElement>().FirstOrDefault();
                if (expectation.Kind == "formula")
                {
                    if (expectation.ExplicitCache is { } expectedCache)
                    {
                        if (cell.GetAttribute("t") != (expectedCache.Type == "n" ? "" : expectedCache.Type) ||
                            cached?.InnerText != expectedCache.Value ||
                            cell.GetElementsByTagName("v", Main).Count != 1)
                            issues.Add(new("G4", "INTENT_CACHE_MISMATCH", $"{group.Key}!{expectation.Address}"));
                    }
                    else if (!expectation.KeepCache && (cached is not null || cell.HasAttribute("t")))
                        issues.Add(new("G4", "INTENT_CACHE_PRESENT", $"{group.Key}!{expectation.Address}"));
                }
                var scalar = cached?.InnerText;
                var type = cell.GetAttribute("t");
                var actual = formula is not null ? ("formula", formula.InnerText) : type switch
                {
                    "s" when int.TryParse(scalar, out var index) && index >= 0 && index < shared.Length => ("text", shared[index]),
                    "inlineStr" => ("inline", cell!.ChildNodes.OfType<XmlElement>().FirstOrDefault(child => child.LocalName == "is" && child.NamespaceURI == Main) is { } inline ? TextValue(inline) : null),
                    "b" => ("boolean", scalar == "1" ? "true" : "false"),
                    "e" => ("error", scalar),
                    _ when cell is null || cell.GetElementsByTagName("v", Main).Count == 0 => ("blank", (string?)null),
                    _ => ("number", scalar)
                };
                var expectedValue = expectation.Kind == "formula" ? expectation.Value?.TrimStart('=') : expectation.Value;
                if (actual.Item1 != expectation.Kind && !(actual.Item1 == "inline" && expectation.Kind == "text") || actual.Item2 != expectedValue)
                    issues.Add(new("G4", "INTENT_MISMATCH", $"{group.Key}!{expectation.Address}: expected {expectation.Kind} {expectedValue}, found {actual.Item1} {actual.Item2}"));
            }
        }
        return issues;
    }

    private static string TextValue(XmlElement container) => string.Concat(container.ChildNodes.OfType<XmlElement>()
        .Where(child => child.NamespaceURI == Main)
        .Select(child => child.LocalName switch
        {
            "t" => child.InnerText,
            "r" => string.Concat(child.ChildNodes.OfType<XmlElement>()
                .Where(text => text.LocalName == "t" && text.NamespaceURI == Main).Select(text => text.InnerText)),
            _ => ""
        }));

    private static XmlDocument Load(ZipArchiveEntry entry)
    {
        var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        document.Load(reader);
        return document;
    }

    private static byte[] Read(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static string? SourcePart(string path)
    {
        if (path.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase)) return "";
        var index = path.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        return index < 0 || !path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ? null :
            path[..index] + "/" + path[(index + "/_rels/".Length)..^5];
    }

    private static string Resolve(string source, string target)
    {
        var decoded = Uri.UnescapeDataString(target.Split('#')[0]);
        var components = (decoded.StartsWith('/') ? decoded.TrimStart('/') : source.Contains('/') ? source[..(source.LastIndexOf('/') + 1)] + decoded : decoded).Split('/');
        var result = new List<string>();
        foreach (var component in components)
        {
            if (component == ".") continue;
            if (component == "..") { if (result.Count == 0) throw new InvalidDataException("Target escapes package"); result.RemoveAt(result.Count - 1); }
            else if (component.Length == 0 || component.Contains('\\')) throw new InvalidDataException("Invalid OPC segment");
            else result.Add(component);
        }
        return string.Join('/', result);
    }

    private static Dictionary<string, (string Type, string Target)> Relations(Dictionary<string, ZipArchiveEntry> entries, string source)
    {
        var slash = source.LastIndexOf('/');
        var path = source == "" ? "_rels/.rels" : (slash < 0 ? "" : source[..(slash + 1)]) + "_rels/" + source[(slash + 1)..] + ".rels";
        return Load(entries[path]).GetElementsByTagName("Relationship", Relationships).OfType<XmlElement>()
            .ToDictionary(item => item.GetAttribute("Id"), item => (item.GetAttribute("Type"), item.GetAttribute("Target")));
    }
}
