using System.IO.Compression;
using System.Xml;

namespace DocLoupe.Excel.Package;

public sealed record PartRelationship(string Id, string Type, string Target, bool External);

public sealed class PackageStore : IDisposable
{
    public const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    public const string ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private readonly ZipArchive _archive;
    private readonly Dictionary<string, ZipArchiveEntry> _entries;
    private readonly Dictionary<string, byte[]> _overlay = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _removed = new(StringComparer.OrdinalIgnoreCase);

    public PackageStore(string path)
    {
        _archive = ZipFile.OpenRead(path);
        try
        {
            _entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in _archive.Entries)
            {
                if (!_entries.TryAdd(entry.FullName, entry))
                    throw new InvalidDataException($"Duplicate OPC part: {entry.FullName}");
            }
            if (!_entries.ContainsKey("[Content_Types].xml") || !_entries.ContainsKey("_rels/.rels"))
                throw new InvalidDataException("Missing OPC root parts");
            var main = ReadRelationships("").SingleOrDefault(relationship => relationship.Type.EndsWith("/officeDocument", StringComparison.Ordinal));
            if (main is null || main.External)
                throw new InvalidDataException("Missing internal officeDocument relationship");
            WorkbookPart = Resolve("", main.Target);
            if (!Contains(WorkbookPart) || !Contains(RelationshipPart(WorkbookPart)))
                throw new InvalidDataException("Missing workbook or workbook relationships");
        }
        catch
        {
            _archive.Dispose();
            throw;
        }
    }

    public string WorkbookPart { get; }
    public IEnumerable<string> Parts => _entries.Keys.Concat(_overlay.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
        .Where(part => !_removed.Contains(part));
    public IEnumerable<string> ChangedParts => _overlay.Keys.Concat(_removed).Distinct(StringComparer.OrdinalIgnoreCase);
    public bool Contains(string part) => !_removed.Contains(part) && (_overlay.ContainsKey(part) || _entries.ContainsKey(part));

    public byte[] Read(string part)
    {
        if (_removed.Contains(part)) throw new FileNotFoundException("Removed OPC part", part);
        if (_overlay.TryGetValue(part, out var bytes)) return bytes;
        return ReadOriginal(part);
    }

    public byte[] ReadOriginal(string part)
    {
        if (!_entries.TryGetValue(part, out var entry)) throw new FileNotFoundException("Missing OPC part", part);
        using var stream = entry.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    public void Set(string part, byte[] content)
    {
        if (part.StartsWith('/') || part.Contains('\\') || part.Split('/').Any(segment => segment is "" or "." or ".."))
            throw new InvalidDataException($"Invalid OPC part: {part}");
        _removed.Remove(part);
        _overlay[part] = content;
    }

    public void Delete(string part)
    {
        if (!Contains(part)) return;
        _overlay.Remove(part);
        _removed.Add(part);
    }

    public static XmlDocument Parse(byte[] content)
    {
        var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
        using var stream = new MemoryStream(content);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        document.Load(reader);
        return document;
    }

    public IReadOnlyList<PartRelationship> ReadRelationships(string source)
    {
        var path = RelationshipPart(source);
        if (!Contains(path)) return [];
        var root = Parse(Read(path)).DocumentElement;
        if (root is null || root.LocalName != "Relationships" || root.NamespaceURI != PackageRelationships)
            throw new InvalidDataException($"Invalid relationships: {path}");
        return root.ChildNodes.OfType<XmlElement>().Where(element => element.LocalName == "Relationship" && element.NamespaceURI == PackageRelationships)
            .Select(element => new PartRelationship(element.GetAttribute("Id"), element.GetAttribute("Type"), element.GetAttribute("Target"),
                element.GetAttribute("TargetMode") == "External")).ToArray();
    }

    public string SheetPart(string name)
    {
        var root = Parse(Read(WorkbookPart)).DocumentElement ?? throw new InvalidDataException("Missing workbook root");
        var sheet = root.GetElementsByTagName("sheet", Main).OfType<XmlElement>()
            .SingleOrDefault(element => element.GetAttribute("name") == name)
            ?? throw new KeyNotFoundException($"Sheet not found: {name}");
        var id = sheet.GetAttribute("id", OfficeRelationships);
        var relationship = ReadRelationships(WorkbookPart).SingleOrDefault(candidate => candidate.Id == id && !candidate.External)
            ?? throw new InvalidDataException($"Unresolved sheet relationship: {id}");
        var part = Resolve(WorkbookPart, relationship.Target);
        if (!Contains(part)) throw new InvalidDataException($"Missing worksheet: {part}");
        return part;
    }

    public IReadOnlyList<string> SheetNames() => Parse(Read(WorkbookPart)).GetElementsByTagName("sheet", Main)
        .OfType<XmlElement>().Select(element => element.GetAttribute("name")).ToArray();

    public static string RelationshipPart(string source)
    {
        if (source == "") return "_rels/.rels";
        var slash = source.LastIndexOf('/');
        return (slash < 0 ? "" : source[..(slash + 1)]) + "_rels/" + source[(slash + 1)..] + ".rels";
    }

    public static string Resolve(string source, string target)
    {
        var clean = Uri.UnescapeDataString(target.Split('#')[0]);
        if (clean.Contains('\\')) throw new InvalidDataException("Backslash in OPC target");
        var combined = clean.StartsWith('/') ? clean[1..] : (source.Contains('/') ? source[..(source.LastIndexOf('/') + 1)] : "") + clean;
        var segments = new List<string>();
        foreach (var segment in combined.Split('/'))
        {
            if (segment == ".") continue;
            if (segment == "..")
            {
                if (segments.Count == 0) throw new InvalidDataException("OPC target escapes package");
                segments.RemoveAt(segments.Count - 1);
            }
            else if (segment.Length == 0) throw new InvalidDataException("Empty OPC path segment");
            else segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    public void Save(string output)
    {
        using var destination = new ZipArchive(File.Create(output), ZipArchiveMode.Create);
        foreach (var entry in _archive.Entries)
        {
            if (_removed.Contains(entry.FullName)) continue;
            var written = destination.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            written.LastWriteTime = entry.LastWriteTime;
            using var target = written.Open();
            if (_overlay.TryGetValue(entry.FullName, out var bytes)) target.Write(bytes);
            else using (var source = entry.Open()) source.CopyTo(target);
        }
        foreach (var (part, bytes) in _overlay.Where(pair => !_entries.ContainsKey(pair.Key)))
        {
            var entry = destination.CreateEntry(part, CompressionLevel.Optimal);
            using var stream = entry.Open();
            stream.Write(bytes);
        }
    }

    public void Dispose() => _archive.Dispose();
}
