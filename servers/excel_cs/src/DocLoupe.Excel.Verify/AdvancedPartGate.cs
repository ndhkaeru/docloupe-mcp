using System.IO.Compression;
using System.Xml;

namespace DocLoupe.Excel.Verify;

public static class AdvancedPartGate
{
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string PackageTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string SharedStringsType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    private const string VbaType = "application/vnd.ms-office.vbaProject";

    public static IReadOnlyList<GateIssue> CheckSignedSource(string source)
    {
        using var zip = ZipFile.OpenRead(source);
        var entries = zip.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        if (entries.Keys.Any(part => part.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)) ||
            ReadRelationships(entries).Values.SelectMany(items => items)
                .Any(item => item.Type.Contains("/digital-signature/", StringComparison.OrdinalIgnoreCase)) ||
            ReadContentTypes(entries).Values.Any(type => type.Contains("digital-signature", StringComparison.OrdinalIgnoreCase)))
            return [new GateIssue("G6", "SIGNED_PACKAGE_UNSUPPORTED", "Editing a signed package would invalidate its signature")];
        return [];
    }

    public static IReadOnlyList<GateIssue> Check(string source, string written, IEnumerable<CellExpectation> expected)
    {
        var issues = new List<GateIssue>();
        issues.AddRange(CheckSignedSource(source));
        if (issues.Count > 0) return issues;
        using var original = ZipFile.OpenRead(source);
        using var staged = ZipFile.OpenRead(written);
        var before = original.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var after = staged.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var oldTypes = ReadContentTypes(before);
        var newTypes = ReadContentTypes(after);
        var oldRelations = ReadRelationships(before);
        var newRelations = ReadRelationships(after);
        var main = Resolve("", oldRelations["_rels/.rels"].Single(item => item.Type.EndsWith("/officeDocument", StringComparison.Ordinal)).Target);
        var workbookRels = RelationshipPart(main);
        var oldWorkbookRelations = oldRelations[workbookRels];
        var oldChain = oldWorkbookRelations.FirstOrDefault(item => item.Type.EndsWith("/calcChain", StringComparison.Ordinal));
        var chain = oldChain is null ? null : Resolve(main, oldChain.Target);
        var hadSharedStrings = oldWorkbookRelations.Any(item => item.Type.EndsWith("/sharedStrings", StringComparison.Ordinal));
        var mutableParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "[Content_Types].xml", main, workbookRels };
        foreach (var relation in oldWorkbookRelations.Where(item => item.Type.EndsWith("/worksheet", StringComparison.Ordinal) ||
                     item.Type.EndsWith("/sharedStrings", StringComparison.Ordinal)))
            mutableParts.Add(Resolve(main, relation.Target));
        var formulaOverwritten = false;
        foreach (var group in expected.Where(item => item.Kind != "formula").GroupBy(item => item.Sheet))
        {
            try
            {
                var existing = P2aGates.ReadCells(source, group.Key, group.Select(item => item.Address));
                formulaOverwritten |= existing.Any(item => item.Formula is not null);
            }
            catch (Exception error) when (error is InvalidDataException or InvalidOperationException or KeyNotFoundException)
            {
                issues.Add(new("G6", "FORMULA_ORIGIN_UNVERIFIED", group.Key));
            }
        }
        if (chain is not null)
        {
            if (formulaOverwritten && after.ContainsKey(chain))
                issues.Add(new("G6", "STALE_CALC_CHAIN", chain));
            if (!formulaOverwritten && !after.ContainsKey(chain))
                issues.Add(new("G6", "CALC_CHAIN_REMOVED", chain));
            if (formulaOverwritten && after.ContainsKey(chain) == false && HasOverride(after, chain))
                issues.Add(new("G6", "STALE_CALC_CHAIN_TYPE", chain));
        }

        foreach (var part in before.Keys)
        {
            if (part.Equals(chain, StringComparison.OrdinalIgnoreCase) && formulaOverwritten && !after.ContainsKey(part))
                continue;
            if (!after.ContainsKey(part)) continue;
            if (ContentType(oldTypes, part) != ContentType(newTypes, part))
                issues.Add(new("G6", "CONTENT_TYPE_CHANGED", part));
        }
        foreach (var (part, entry) in before)
        {
            if (mutableParts.Contains(part) ||
                part.Equals(chain, StringComparison.OrdinalIgnoreCase) && formulaOverwritten)
                continue;
            if (!after.TryGetValue(part, out var current))
            {
                issues.Add(new("G6", "PROTECTED_PART_MISSING", part));
                continue;
            }
            if (!EqualBytes(entry, current)) issues.Add(new("G6", "PROTECTED_PART_CHANGED", part));
            if (part.EndsWith("/vbaProject.bin", StringComparison.OrdinalIgnoreCase) &&
                ContentType(newTypes, part) != VbaType)
                issues.Add(new("G6", "VBA_CONTENT_TYPE", part));
        }
        foreach (var (part, entry) in after.Where(item => !before.ContainsKey(item.Key)))
        {
            if (hadSharedStrings || ContentType(newTypes, part) != SharedStringsType ||
                !newRelations.GetValueOrDefault(workbookRels, []).Any(item =>
                    item.Type.EndsWith("/sharedStrings", StringComparison.Ordinal) &&
                    Resolve(main, item.Target).Equals(part, StringComparison.OrdinalIgnoreCase)))
                issues.Add(new("G6", "UNEXPECTED_PART_ADDED", part));
        }
        foreach (var (part, oldItems) in oldRelations)
        {
            if (!newRelations.TryGetValue(part, out var current))
            {
                issues.Add(new("G6", "RELATIONSHIPS_REMOVED", part));
                continue;
            }
            foreach (var item in oldItems.Except(current))
                if (!(part.Equals(workbookRels, StringComparison.OrdinalIgnoreCase) && formulaOverwritten && item == oldChain))
                    issues.Add(new("G6", "RELATIONSHIP_CHANGED", part + ": " + item.Id));
            foreach (var item in current.Except(oldItems))
                if (!(part.Equals(workbookRels, StringComparison.OrdinalIgnoreCase) &&
                    item.Type.EndsWith("/sharedStrings", StringComparison.Ordinal) && !hadSharedStrings &&
                    !before.ContainsKey(Resolve(main, item.Target)) && after.ContainsKey(Resolve(main, item.Target))))
                    issues.Add(new("G6", "RELATIONSHIP_ADDED", part + ": " + item.Id));
        }
        foreach (var part in newRelations.Keys.Except(oldRelations.Keys, StringComparer.OrdinalIgnoreCase))
            issues.Add(new("G6", "RELATIONSHIPS_ADDED", part));
        return issues;
    }

    private static bool EqualBytes(ZipArchiveEntry oldEntry, ZipArchiveEntry newEntry)
    {
        if (oldEntry.Length != newEntry.Length) return false;
        using var original = oldEntry.Open();
        using var current = newEntry.Open();
        var left = new byte[8192];
        var right = new byte[8192];
        int read;
        while ((read = original.Read(left)) != 0)
        {
            current.ReadExactly(right.AsSpan(0, read));
            if (!left.AsSpan(0, read).SequenceEqual(right.AsSpan(0, read))) return false;
        }
        return current.ReadByte() == -1;
    }

    private static bool HasOverride(Dictionary<string, ZipArchiveEntry> entries, string part) =>
        Load(entries["[Content_Types].xml"]).GetElementsByTagName("Override", PackageTypes).OfType<XmlElement>()
            .Any(item => item.GetAttribute("PartName").TrimStart('/').Equals(part, StringComparison.OrdinalIgnoreCase));

    private static Dictionary<string, string> ReadContentTypes(Dictionary<string, ZipArchiveEntry> entries)
    {
        var types = Load(entries["[Content_Types].xml"]);
        var defaults = types.GetElementsByTagName("Default", PackageTypes).OfType<XmlElement>()
            .ToDictionary(item => "." + item.GetAttribute("Extension"), item => item.GetAttribute("ContentType"), StringComparer.OrdinalIgnoreCase);
        var overrides = types.GetElementsByTagName("Override", PackageTypes).OfType<XmlElement>()
            .ToDictionary(item => item.GetAttribute("PartName").TrimStart('/'), item => item.GetAttribute("ContentType"), StringComparer.OrdinalIgnoreCase);
        return entries.Keys.ToDictionary(part => part, part => overrides.GetValueOrDefault(part) ??
            defaults.GetValueOrDefault(Path.GetExtension(part)) ?? "", StringComparer.OrdinalIgnoreCase);
    }

    private static string? ContentType(Dictionary<string, string> types, string part) =>
        types.GetValueOrDefault(part);

    private static Dictionary<string, List<Relationship>> ReadRelationships(Dictionary<string, ZipArchiveEntry> entries) =>
        entries.Where(item => item.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(item => item.Key, item => Load(item.Value)
                .GetElementsByTagName("Relationship", PackageRelationships).OfType<XmlElement>()
                .Select(element => new Relationship(element.GetAttribute("Id"), element.GetAttribute("Type"),
                    element.GetAttribute("Target"), element.GetAttribute("TargetMode"))).ToList(), StringComparer.OrdinalIgnoreCase);

    private static XmlDocument Load(ZipArchiveEntry entry)
    {
        var document = new XmlDocument { XmlResolver = null };
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        document.Load(reader);
        return document;
    }

    private static string RelationshipPart(string source)
    {
        var position = source.LastIndexOf('/');
        return (position < 0 ? "" : source[..(position + 1)]) + "_rels/" + source[(position + 1)..] + ".rels";
    }

    private static string Resolve(string source, string target)
    {
        var decoded = Uri.UnescapeDataString(target.Split('#')[0]);
        var parts = (decoded.StartsWith('/') ? decoded.TrimStart('/') :
            source.Contains('/') ? source[..(source.LastIndexOf('/') + 1)] + decoded : decoded).Split('/');
        var path = new List<string>();
        foreach (var part in parts)
        {
            if (part == ".") continue;
            if (part == "..") { if (path.Count == 0) throw new InvalidDataException("Relationship escapes package"); path.RemoveAt(path.Count - 1); }
            else if (part.Length == 0 || part.Contains('\\')) throw new InvalidDataException("Invalid relationship target");
            else path.Add(part);
        }
        return string.Join('/', path);
    }

    private sealed record Relationship(string Id, string Type, string Target, string TargetMode);
}
