using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocLoupe.Excel.Model;

namespace DocLoupe.Excel.Verify;

public static partial class P2aGates
{
    public static IReadOnlyList<GateIssue> CheckSemanticPreservation(string source, string written,
        IEnumerable<CellExpectation> intended, IEnumerable<DeclaredByteSpan>? declared = null)
    {
        try
        {
            using var oldZip = ZipFile.OpenRead(source);
            using var newZip = ZipFile.OpenRead(written);
            var oldEntries = oldZip.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
            var newEntries = newZip.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
            var oldWorkbook = SemanticWorkbook.Read(oldEntries);
            var newWorkbook = SemanticWorkbook.Read(newEntries);
            var expectations = intended.ToArray();
            var targets = expectations.GroupBy(item => item.Sheet, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Select(item => item.Address).ToHashSet(StringComparer.Ordinal), StringComparer.Ordinal);
            var issues = new List<GateIssue>();
            var handled = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                oldWorkbook.Part, newWorkbook.Part, oldWorkbook.RelationshipPart, newWorkbook.RelationshipPart,
                "[Content_Types].xml"
            };
            var formulaRemoved = false;
            foreach (var name in oldWorkbook.Sheets.Keys.Union(newWorkbook.Sheets.Keys, StringComparer.Ordinal))
            {
                if (!oldWorkbook.Sheets.TryGetValue(name, out var oldPart) ||
                    !newWorkbook.Sheets.TryGetValue(name, out var newPart) ||
                    !oldPart.Equals(newPart, StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add(new("G5", "UNDECLARED_SHEET_CHANGE", name));
                    continue;
                }
                handled.Add(oldPart);
                var original = Load(oldEntries[oldPart]);
                var candidate = Load(newEntries[newPart]);
                var selected = targets.GetValueOrDefault(name) ?? [];
                Dictionary<string, XmlElement> SelectedCells(XmlDocument document)
                {
                    var groups = document.GetElementsByTagName("c", Main).OfType<XmlElement>()
                        .Where(cell => selected.Contains(cell.GetAttribute("r")))
                        .GroupBy(cell => cell.GetAttribute("r"), StringComparer.Ordinal);
                    var cells = new Dictionary<string, XmlElement>(StringComparer.Ordinal);
                    foreach (var group in groups)
                    {
                        if (group.Skip(1).Any()) issues.Add(new("G5", "AMBIGUOUS_TARGET", $"{name}!{group.Key}"));
                        cells[group.Key] = group.First();
                    }
                    return cells;
                }
                var beforeTargets = SelectedCells(original);
                var afterTargets = SelectedCells(candidate);
                foreach (var address in selected)
                {
                    beforeTargets.TryGetValue(address, out var oldCell);
                    afterTargets.TryGetValue(address, out var newCell);
                    if (oldCell?.ChildNodes.OfType<XmlElement>().Any(child => child.LocalName == "f" && child.NamespaceURI == Main) == true &&
                        newCell?.ChildNodes.OfType<XmlElement>().Any(child => child.LocalName == "f" && child.NamespaceURI == Main) != true)
                        formulaRemoved = true;
                }
                var originalRows = original.GetElementsByTagName("row", Main).OfType<XmlElement>()
                    .Select(row => row.GetAttribute("r")).ToHashSet(StringComparer.Ordinal);
                StripSelectedCells(original, selected, originalRows);
                StripSelectedCells(candidate, selected, originalRows);
                if (!Equivalent(original.DocumentElement, candidate.DocumentElement))
                {
                    var oldCells = original.GetElementsByTagName("c", Main).OfType<XmlElement>()
                        .ToDictionary(cell => cell.GetAttribute("r"), StringComparer.Ordinal);
                    var newCells = candidate.GetElementsByTagName("c", Main).OfType<XmlElement>()
                        .ToDictionary(cell => cell.GetAttribute("r"), StringComparer.Ordinal);
                    var differingCell = oldCells.Keys.Union(newCells.Keys, StringComparer.Ordinal)
                        .FirstOrDefault(address => !Equivalent(oldCells.GetValueOrDefault(address),
                            newCells.GetValueOrDefault(address)));
                    issues.Add(new("G5", "UNDECLARED_SHEET_CHANGE", $"{name}!{differingCell ?? "sheetData"}"));
                }
            }
            if (declared is not null)
                CheckDeclarationScope(oldEntries, newEntries, oldWorkbook, newWorkbook, targets,
                    formulaRemoved, declared, issues);
            var oldShared = oldWorkbook.SharedPart;
            var newShared = newWorkbook.SharedPart;
            if (oldShared is not null) handled.Add(oldShared);
            if (newShared is not null) handled.Add(newShared);
            var addedShared = oldShared is null && newShared is not null;
            if (oldShared is not null && newShared is null || oldShared is not null && newShared is not null &&
                !oldShared.Equals(newShared, StringComparison.OrdinalIgnoreCase))
                issues.Add(new("G5", "UNDECLARED_SHARED_STRING_CHANGE", "sharedStrings relationship"));
            else if (newShared is not null && newEntries.TryGetValue(newShared, out var sharedEntry))
                CheckSharedStrings(oldShared is null ? null : Load(oldEntries[oldShared]), Load(sharedEntry),
                    oldWorkbook, oldEntries, newWorkbook, newEntries, targets, expectations, issues);
            if (!EquivalentWorkbook(oldWorkbook.Document, newWorkbook.Document, targets.Count > 0))
                issues.Add(new("G5", "UNDECLARED_WORKBOOK_CHANGE", oldWorkbook.Part));
            var removedChain = oldWorkbook.CalcChainPart is not null && newWorkbook.CalcChainPart is null && formulaRemoved;
            if (oldWorkbook.CalcChainPart is not null) handled.Add(oldWorkbook.CalcChainPart);
            if (newWorkbook.CalcChainPart is not null) handled.Add(newWorkbook.CalcChainPart);
            if (oldWorkbook.CalcChainPart != newWorkbook.CalcChainPart && !removedChain ||
                oldWorkbook.CalcChainPart is not null && newWorkbook.CalcChainPart is not null &&
                !Read(oldEntries[oldWorkbook.CalcChainPart]).AsSpan().SequenceEqual(Read(newEntries[newWorkbook.CalcChainPart])))
                issues.Add(new("G5", "UNDECLARED_CALC_CHAIN_CHANGE", oldWorkbook.CalcChainPart ?? "calcChain"));
            CompareRelationships(oldEntries[oldWorkbook.RelationshipPart], newEntries[newWorkbook.RelationshipPart],
                oldWorkbook.Part, newShared, oldWorkbook.CalcChainPart, addedShared, removedChain, issues);
            CompareContentTypes(oldEntries["[Content_Types].xml"], newEntries["[Content_Types].xml"],
                newShared, oldWorkbook.CalcChainPart, addedShared, removedChain, issues);
            foreach (var name in oldEntries.Keys.Union(newEntries.Keys, StringComparer.OrdinalIgnoreCase))
            {
                if (handled.Contains(name)) continue;
                if (!oldEntries.TryGetValue(name, out var oldEntry) || !newEntries.TryGetValue(name, out var newEntry))
                {
                    issues.Add(new("G5", "UNDECLARED_PART_CHANGE", name));
                    continue;
                }
                if (!Read(oldEntry).AsSpan().SequenceEqual(Read(newEntry)))
                    issues.Add(new("G5", "UNDECLARED_PART_CHANGE", name));
            }
            if (removedChain && oldWorkbook.CalcChainPart is not null && newEntries.ContainsKey(oldWorkbook.CalcChainPart))
                issues.Add(new("G5", "UNDECLARED_PART_CHANGE", oldWorkbook.CalcChainPart));
            if (newShared is not null && oldShared is null && !newEntries.ContainsKey(newShared))
                issues.Add(new("G5", "UNDECLARED_SHARED_STRING_CHANGE", "missing sharedStrings part"));
            return issues;
        }
        catch (Exception exception) when (exception is InvalidDataException or XmlException or KeyNotFoundException or ArgumentException or InvalidOperationException)
        {
            return [new("G5", "SEMANTIC_CHECK_FAILED", exception.Message)];
        }
    }

    private sealed record SemanticWorkbook(string Part, string RelationshipPart, XmlDocument Document,
        Dictionary<string, string> Sheets, string? SharedPart, string? CalcChainPart)
    {
        public static SemanticWorkbook Read(Dictionary<string, ZipArchiveEntry> entries)
        {
            var root = Relations(entries, "");
            var main = Resolve("", root.Values.Single(relationship => relationship.Type.EndsWith("/officeDocument", StringComparison.Ordinal)).Target);
            var document = Load(entries[main]);
            var relationships = Relations(entries, main);
            var sheets = document.GetElementsByTagName("sheet", Main).OfType<XmlElement>()
                .ToDictionary(sheet => sheet.GetAttribute("name"), sheet =>
                    Resolve(main, relationships[sheet.GetAttribute("id", Office)].Target), StringComparer.Ordinal);
            string? Target(string type) => relationships.Values.FirstOrDefault(relation =>
                relation.Type.EndsWith("/" + type, StringComparison.Ordinal)) is { Type: not null } relation
                    ? Resolve(main, relation.Target) : null;
            var slash = main.LastIndexOf('/');
            var rels = (slash < 0 ? "" : main[..(slash + 1)]) + "_rels/" + main[(slash + 1)..] + ".rels";
            return new(main, rels, document, sheets, Target("sharedStrings"), Target("calcChain"));
        }
    }

    private static void CheckDeclarationScope(Dictionary<string, ZipArchiveEntry> entries,
        Dictionary<string, ZipArchiveEntry> writtenEntries,
        SemanticWorkbook workbook, SemanticWorkbook writtenWorkbook,
        Dictionary<string, HashSet<string>> targets, bool formulaRemoved,
        IEnumerable<DeclaredByteSpan> declared, List<GateIssue> issues)
    {
        var byPart = workbook.Sheets.ToDictionary(sheet => sheet.Value, sheet => sheet.Key,
            StringComparer.OrdinalIgnoreCase);
        foreach (var group in declared.GroupBy(edit => edit.Part, StringComparer.OrdinalIgnoreCase))
        {
            if (!entries.TryGetValue(group.Key, out var entry))
            {
                if (group.Any()) issues.Add(new("G5", "DECLARATION_OUTSIDE_INTENT", group.Key));
                continue;
            }
            var source = Read(entry);
            var text = new UTF8Encoding(false, true).GetString(source);
            if (!byPart.TryGetValue(group.Key, out var sheet))
            {
                var writtenText = new UTF8Encoding(false, true).GetString(Read(writtenEntries[group.Key]));
                CheckMetadataScope(group.Key, text, writtenText, workbook, writtenWorkbook,
                    formulaRemoved, group, issues);
                continue;
            }
            var allowed = targets.GetValueOrDefault(sheet) ?? [];
            var cellSpans = new List<(int Start, int End, string Address)>();
            foreach (Match opening in Regex.Matches(text,
                @"<(?<prefix>[A-Za-z_][\w.-]*:)?c(?=\s|/|>)(?<attributes>(?:""[^""]*""|'[^']*'|[^'""<>])*)>",
                RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5)))
            {
                var address = Regex.Match(opening.Groups["attributes"].Value,
                    @"(?<![\w:])r\s*=\s*(['""])(?<value>.*?)\1", RegexOptions.CultureInvariant,
                    TimeSpan.FromSeconds(5));
                if (!address.Success) continue;
                var end = opening.Index + opening.Length;
                if (!opening.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal))
                {
                    var closing = Regex.Match(text[end..],
                        @"</" + Regex.Escape(opening.Groups["prefix"].Value) + @"c\s*>",
                        RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));
                    if (!closing.Success) throw new InvalidDataException("Unclosed worksheet cell");
                    end += closing.Index + closing.Length;
                }
                cellSpans.Add((Encoding.UTF8.GetByteCount(text.AsSpan(0, opening.Index)),
                    Encoding.UTF8.GetByteCount(text.AsSpan(0, end)), address.Groups["value"].Value));
            }
            var insertionOffsets = new HashSet<int>();
            var rows = TagSpans(text, "row");
            var sheetData = TagSpans(text, "sheetData").SingleOrDefault();
            foreach (var address in allowed.Where(address => cellSpans.All(cell => cell.Address != address)))
            {
                var target = CellAddress.Parse(address);
                var row = rows.SingleOrDefault(candidate => RawAttribute(candidate, "r") ==
                    target.Row.ToString(CultureInfo.InvariantCulture));
                if (row is not null)
                {
                    var end = RowEndTagOffset(text, row);
                    var nextCell = cellSpans.Where(cell => cell.Start >= row.End && cell.Start < end &&
                            CellAddress.Parse(cell.Address).Column > target.Column)
                        .OrderBy(cell => cell.Start).FirstOrDefault();
                    insertionOffsets.Add(nextCell.Address is null ? end : nextCell.Start);
                }
                else if (sheetData is not null)
                {
                    var nextRow = rows.Where(candidate => int.Parse(RawAttribute(candidate, "r"),
                            CultureInfo.InvariantCulture) > target.Row)
                        .OrderBy(candidate => candidate.Start).FirstOrDefault();
                    insertionOffsets.Add(nextRow?.Start ?? ClosingTagOffset(text, sheetData));
                }
                else
                    insertionOffsets.Add(ClosingTagOffset(text, TagSpans(text, "worksheet").Single()));
            }
            foreach (var edit in group)
            {
                var permitted = edit.Start == edit.End
                    ? insertionOffsets.Contains(edit.Start)
                    : cellSpans.Any(cell => allowed.Contains(cell.Address) &&
                        edit.Start >= cell.Start && edit.End <= cell.End);
                if (!permitted)
                    issues.Add(new("G5", "DECLARATION_OUTSIDE_INTENT", $"{sheet}: {edit.Start}-{edit.End}"));
            }
        }
    }

    private sealed record RawTag(int Start, int End, string Name, string Attributes,
        int CharStart, int CharEnd);

    private static IReadOnlyList<RawTag> TagSpans(string text, string localName)
    {
        var pattern = @"<(?<prefix>[A-Za-z_][\w.-]*:)?" + Regex.Escape(localName) +
            @"(?=\s|/|>)(?<attributes>(?:""[^""]*""|'[^']*'|[^'""<>])*)>";
        return Regex.Matches(text, pattern, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5))
            .Cast<Match>()
            .Select(match => new RawTag(Encoding.UTF8.GetByteCount(text.AsSpan(0, match.Index)),
                Encoding.UTF8.GetByteCount(text.AsSpan(0, match.Index + match.Length)),
                match.Groups["prefix"].Value + localName, match.Groups["attributes"].Value,
                match.Index, match.Index + match.Length))
            .ToArray();
    }

    private static RawTag CompleteEmptyTag(string text, RawTag tag)
    {
        if (text.AsSpan(tag.CharStart, tag.CharEnd - tag.CharStart).TrimEnd().EndsWith("/>".AsSpan(), StringComparison.Ordinal))
            return tag;
        var closing = new Regex(@"\G</" + Regex.Escape(tag.Name) + @"\s*>",
            RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5)).Match(text, tag.CharEnd);
        return closing.Success
            ? tag with { End = Encoding.UTF8.GetByteCount(text.AsSpan(0, closing.Index + closing.Length)) }
            : tag;
    }

    private static string RawAttribute(RawTag tag, string name)
    {
        var attributes = tag.Attributes;
        for (var offset = 0; offset < attributes.Length;)
        {
            while (offset < attributes.Length && char.IsWhiteSpace(attributes[offset])) offset++;
            if (offset == attributes.Length || attributes[offset] == '/') break;
            var start = offset;
            while (offset < attributes.Length && !char.IsWhiteSpace(attributes[offset]) &&
                   attributes[offset] is not ('=' or '/')) offset++;
            var attributeName = attributes[start..offset];
            while (offset < attributes.Length && char.IsWhiteSpace(attributes[offset])) offset++;
            if (attributeName.Length == 0 || offset == attributes.Length || attributes[offset++] != '=')
                throw new InvalidDataException("Malformed package attribute");
            while (offset < attributes.Length && char.IsWhiteSpace(attributes[offset])) offset++;
            if (offset == attributes.Length || attributes[offset] is not ('\'' or '"'))
                throw new InvalidDataException("Malformed package attribute value");
            var quote = attributes[offset++];
            var valueStart = offset;
            while (offset < attributes.Length && attributes[offset] != quote) offset++;
            if (offset == attributes.Length) throw new InvalidDataException("Unterminated package attribute");
            var value = attributes[valueStart..offset++];
            if (attributeName == name) return value;
        }
        return "";
    }

    private static int ClosingTagOffset(string text, RawTag root)
    {
        var index = text.LastIndexOf("</" + root.Name, StringComparison.Ordinal);
        if (index < 0) throw new InvalidDataException("Missing package XML closing tag");
        return Encoding.UTF8.GetByteCount(text.AsSpan(0, index));
    }

    private static int RowEndTagOffset(string text, RawTag row)
    {
        var index = text.IndexOf("</" + row.Name, row.CharEnd, StringComparison.Ordinal);
        if (index < 0) throw new InvalidDataException("Missing worksheet row closing tag");
        return Encoding.UTF8.GetByteCount(text.AsSpan(0, index));
    }

    private static bool PermittedStartTagChange(string original, string written,
        string localName, IReadOnlyList<string> allowedAttributes)
    {
        var before = TagSpans(original, localName);
        var after = TagSpans(written, localName);
        if (before.Count != 1 || after.Count != 1) return false;
        string WithoutAllowedAttributes(string text, RawTag tag)
        {
            var markup = text[tag.CharStart..tag.CharEnd];
            var prefixEnd = 1 + tag.Name.Length;
            var result = new StringBuilder(markup[..prefixEnd]);
            for (var offset = prefixEnd; offset < markup.Length;)
            {
                var whitespaceStart = offset;
                while (offset < markup.Length && char.IsWhiteSpace(markup[offset])) offset++;
                if (offset == markup.Length || markup[offset] is '/' or '>')
                {
                    result.Append(markup.AsSpan(whitespaceStart));
                    break;
                }
                var nameStart = offset;
                while (offset < markup.Length && !char.IsWhiteSpace(markup[offset]) &&
                       markup[offset] is not ('=' or '/' or '>')) offset++;
                var name = markup[nameStart..offset];
                while (offset < markup.Length && char.IsWhiteSpace(markup[offset])) offset++;
                if (name.Length == 0 || offset >= markup.Length || markup[offset++] != '=')
                    throw new InvalidDataException("Malformed package attribute");
                while (offset < markup.Length && char.IsWhiteSpace(markup[offset])) offset++;
                if (offset >= markup.Length || markup[offset] is not ('\'' or '"'))
                    throw new InvalidDataException("Malformed package attribute value");
                var quote = markup[offset++];
                while (offset < markup.Length && markup[offset] != quote) offset++;
                if (offset == markup.Length) throw new InvalidDataException("Unterminated package attribute");
                offset++;
                if (!allowedAttributes.Contains(name, StringComparer.Ordinal))
                    result.Append(markup.AsSpan(whitespaceStart, offset - whitespaceStart));
            }
            return result.ToString();
        }
        return WithoutAllowedAttributes(original, before[0]) == WithoutAllowedAttributes(written, after[0]);
    }

    private static void CheckMetadataScope(string part, string text, string writtenText,
        SemanticWorkbook workbook, SemanticWorkbook writtenWorkbook, bool formulaRemoved,
        IEnumerable<DeclaredByteSpan> declared, List<GateIssue> issues)
    {
        var allowed = new List<RawTag>();
        var insertionOffsets = new HashSet<int>();
        var insertCalc = false;
        var removeChain = formulaRemoved && workbook.CalcChainPart is not null &&
            writtenWorkbook.CalcChainPart is null;
        var addShared = workbook.SharedPart is null && writtenWorkbook.SharedPart is not null;
        if (part.Equals(workbook.Part, StringComparison.OrdinalIgnoreCase))
        {
            var calc = TagSpans(text, "calcPr");
            if (calc.Count > 1) throw new InvalidDataException("Ambiguous calculation properties");
            if (calc.Count == 1)
            {
                allowed.Add(calc[0]);
                if (!PermittedStartTagChange(text, writtenText, "calcPr", ["fullCalcOnLoad"]))
                    issues.Add(new("G5", "DECLARATION_OUTSIDE_INTENT", part + ": calcPr attributes"));
            }
            else insertCalc = true;
        }
        else if (part.Equals(workbook.SharedPart, StringComparison.OrdinalIgnoreCase))
        {
            var table = TagSpans(text, "sst").Single();
            allowed.Add(table);
            if (!PermittedStartTagChange(text, writtenText, "sst", ["count", "uniqueCount"]))
                issues.Add(new("G5", "DECLARATION_OUTSIDE_INTENT", part + ": sst attributes"));
            insertionOffsets.Add(ClosingTagOffset(text, table));
        }
        else if (part.Equals(workbook.RelationshipPart, StringComparison.OrdinalIgnoreCase))
        {
            if (removeChain)
                allowed.AddRange(TagSpans(text, "Relationship").Where(tag =>
                    RawAttribute(tag, "Type") == Office + "/calcChain" &&
                    Resolve(workbook.Part, RawAttribute(tag, "Target"))
                        .Equals(workbook.CalcChainPart, StringComparison.OrdinalIgnoreCase))
                    .Select(tag => CompleteEmptyTag(text, tag)));
            if (addShared)
                insertionOffsets.Add(ClosingTagOffset(text, TagSpans(text, "Relationships").Single()));
        }
        else if (part.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
        {
            if (removeChain)
                allowed.AddRange(TagSpans(text, "Override").Where(tag =>
                    RawAttribute(tag, "PartName").TrimStart('/')
                        .Equals(workbook.CalcChainPart, StringComparison.OrdinalIgnoreCase))
                    .Select(tag => CompleteEmptyTag(text, tag)));
            if (addShared)
                insertionOffsets.Add(ClosingTagOffset(text, TagSpans(text, "Types").Single()));
        }
        foreach (var edit in declared)
        {
            var permitted = edit.Start == edit.End
                ? insertCalc || insertionOffsets.Contains(edit.Start)
                : allowed.Any(tag => edit.Start >= tag.Start && edit.End <= tag.End);
            if (!permitted)
                issues.Add(new("G5", "DECLARATION_OUTSIDE_INTENT", $"{part}: {edit.Start}-{edit.End}"));
        }
    }

    private static XmlElement? FindCell(XmlDocument document, string address) =>
        document.GetElementsByTagName("c", Main).OfType<XmlElement>()
            .FirstOrDefault(cell => cell.GetAttribute("r") == address);

    private static void StripSelectedCells(XmlDocument document, HashSet<string> selected, HashSet<string> originalRows)
    {
        foreach (var row in document.GetElementsByTagName("row", Main).OfType<XmlElement>().ToArray())
        {
            foreach (var cell in row.ChildNodes.OfType<XmlElement>().Where(cell =>
                cell.LocalName == "c" && cell.NamespaceURI == Main && selected.Contains(cell.GetAttribute("r"))).ToArray())
                row.RemoveChild(cell);
            if (!originalRows.Contains(row.GetAttribute("r")) && row.Attributes.Count == 1 &&
                row.HasAttribute("r") && !row.ChildNodes.OfType<XmlElement>().Any() &&
                string.IsNullOrWhiteSpace(row.InnerText))
                row.ParentNode!.RemoveChild(row);
        }
    }

    private static bool Equivalent(XmlNode? original, XmlNode? written)
    {
        if (original is null || written is null) return original is null && written is null;
        var before = original.CloneNode(true);
        var after = written.CloneNode(true);
        RemoveFormattingWhitespace(before);
        RemoveFormattingWhitespace(after);
        return before.OuterXml == after.OuterXml;
    }

    private static void RemoveFormattingWhitespace(XmlNode node)
    {
        if (node is XmlElement element && element.NamespaceURI == Main && element.LocalName is
            ("worksheet" or "sheetData" or "row" or "workbook" or "sheets" or "sst" or "calcPr"))
            foreach (var child in element.ChildNodes.OfType<XmlCharacterData>().Where(child =>
                child.NodeType is XmlNodeType.Text or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace &&
                string.IsNullOrWhiteSpace(child.Data)).ToArray())
                element.RemoveChild(child);
        foreach (XmlNode child in node.ChildNodes) RemoveFormattingWhitespace(child);
    }

    private static bool EquivalentWorkbook(XmlDocument original, XmlDocument written, bool hasIntent)
    {
        var before = (XmlDocument)original.CloneNode(true);
        var after = (XmlDocument)written.CloneNode(true);
        if (hasIntent)
        {
            foreach (var workbook in new[] { before, after })
            {
                var calc = workbook.DocumentElement!.ChildNodes.OfType<XmlElement>()
                    .FirstOrDefault(child => child.LocalName == "calcPr" && child.NamespaceURI == Main);
                calc?.RemoveAttribute("fullCalcOnLoad");
                if (calc is not null && calc.Attributes.Count == 0 && calc.ChildNodes.Count == 0 &&
                    before.DocumentElement!.ChildNodes.OfType<XmlElement>().All(child => child.LocalName != "calcPr"))
                    calc.ParentNode!.RemoveChild(calc);
            }
        }
        return Equivalent(before.DocumentElement, after.DocumentElement);
    }

    private static void CheckSharedStrings(XmlDocument? original, XmlDocument written,
        SemanticWorkbook oldWorkbook, Dictionary<string, ZipArchiveEntry> oldEntries,
        SemanticWorkbook workbook, Dictionary<string, ZipArchiveEntry> entries,
        Dictionary<string, HashSet<string>> targets,
        IReadOnlyList<CellExpectation> expectations, List<GateIssue> issues)
    {
        var beforeItems = original?.DocumentElement?.ChildNodes.OfType<XmlElement>()
            .Where(item => item.LocalName == "si" && item.NamespaceURI == Main).ToArray() ?? [];
        var afterItems = written.DocumentElement?.ChildNodes.OfType<XmlElement>()
            .Where(item => item.LocalName == "si" && item.NamespaceURI == Main).ToArray() ?? [];
        if (beforeItems.Length > afterItems.Length || beforeItems.Where((item, index) =>
            !Equivalent(item, afterItems[index])).Any())
            issues.Add(new("G5", "UNDECLARED_SHARED_STRING_CHANGE", "existing <si> changed"));
        var allowedIndices = new HashSet<int>();
        foreach (var (sheet, addresses) in targets)
        {
            if (!workbook.Sheets.TryGetValue(sheet, out var part)) continue;
            var document = Load(entries[part]);
            foreach (var address in addresses)
            {
                var cell = FindCell(document, address);
                if (cell?.GetAttribute("t") != "s") continue;
                var value = cell.ChildNodes.OfType<XmlElement>()
                    .FirstOrDefault(child => child.LocalName == "v" && child.NamespaceURI == Main)?.InnerText;
                if (int.TryParse(value, out var index)) allowedIndices.Add(index);
            }
        }
        for (var index = beforeItems.Length; index < afterItems.Length; index++)
        {
            if (!allowedIndices.Contains(index))
                issues.Add(new("G5", "UNDECLARED_SHARED_STRING_CHANGE", $"orphan <si> index {index}"));
            var item = afterItems[index];
            var children = item.ChildNodes.OfType<XmlElement>().ToArray();
            var value = children.Length == 1 ? children[0] : null;
            var preserveSpace = value is { InnerText.Length: > 0 } &&
                (char.IsWhiteSpace(value.InnerText[0]) || char.IsWhiteSpace(value.InnerText[^1]));
            var expectedPrefix = written.DocumentElement!.Prefix;
            var generatedNamespace = item.Attributes.OfType<XmlAttribute>().Where(attribute =>
                attribute.NamespaceURI == "http://www.w3.org/2000/xmlns/" &&
                expectedPrefix.Length > 0 && attribute.LocalName == expectedPrefix && attribute.Value == Main).ToArray();
            var spaceAttributes = value?.Attributes.OfType<XmlAttribute>().Where(attribute =>
                attribute.NamespaceURI == "http://www.w3.org/XML/1998/namespace" &&
                attribute.LocalName == "space" && attribute.Value == "preserve").ToArray() ?? [];
            if (HasUnmodeledNode(item) || item.ChildNodes.Count != 1 ||
                item.Prefix != expectedPrefix || item.Attributes.Count != generatedNamespace.Length ||
                value is null || value.LocalName != "t" || value.NamespaceURI != Main ||
                value.Prefix != expectedPrefix || value.Attributes.Count != spaceAttributes.Length ||
                preserveSpace != (spaceAttributes.Length == 1) ||
                !expectations.Any(expectation => expectation.Kind == "text" && expectation.Value == value.InnerText))
                issues.Add(new("G5", "UNDECLARED_SHARED_STRING_CHANGE", $"unexpected <si> index {index}"));
        }
        var oldReferences = 0;
        var newReferences = 0;
        foreach (var (sheet, addresses) in targets)
        {
            if (!workbook.Sheets.TryGetValue(sheet, out var part)) continue;
            var newSheet = Load(entries[part]);
            newReferences += addresses.Count(address => FindCell(newSheet, address)?.GetAttribute("t") == "s");
            if (oldWorkbook.Sheets.TryGetValue(sheet, out var oldPart))
            {
                var oldSheet = Load(oldEntries[oldPart]);
                oldReferences += addresses.Count(address => FindCell(oldSheet, address)?.GetAttribute("t") == "s");
            }
        }
        if (original is not null)
        {
            var oldCount = original.DocumentElement!.GetAttribute("count");
            var newCount = written.DocumentElement!.GetAttribute("count");
            var oldUnique = original.DocumentElement.GetAttribute("uniqueCount");
            var newUnique = written.DocumentElement.GetAttribute("uniqueCount");
            if (original.DocumentElement.HasAttribute("count") != written.DocumentElement.HasAttribute("count") ||
                original.DocumentElement.HasAttribute("uniqueCount") != written.DocumentElement.HasAttribute("uniqueCount") ||
                oldCount.Length > 0 && (!int.TryParse(oldCount, NumberStyles.None, CultureInfo.InvariantCulture, out var count) ||
                    !int.TryParse(newCount, NumberStyles.None, CultureInfo.InvariantCulture, out var writtenCount) ||
                    writtenCount != count + newReferences - oldReferences) ||
                oldUnique.Length > 0 && (!int.TryParse(oldUnique, NumberStyles.None, CultureInfo.InvariantCulture, out var unique) ||
                    !int.TryParse(newUnique, NumberStyles.None, CultureInfo.InvariantCulture, out var writtenUnique) ||
                    writtenUnique != unique + afterItems.Length - beforeItems.Length))
                issues.Add(new("G5", "UNDECLARED_SHARED_STRING_CHANGE", "sharedStrings counts"));
            foreach (var document in new[] { original, written })
            {
                document.DocumentElement!.RemoveAttribute("count");
                document.DocumentElement.RemoveAttribute("uniqueCount");
                foreach (var item in document.DocumentElement.ChildNodes.OfType<XmlElement>()
                    .Where(item => item.LocalName == "si" && item.NamespaceURI == Main).ToArray())
                    document.DocumentElement.RemoveChild(item);
            }
            if (!Equivalent(original.DocumentElement, written.DocumentElement))
                issues.Add(new("G5", "UNDECLARED_SHARED_STRING_CHANGE", "sharedStrings metadata"));
        }
        else if (written.DocumentElement?.LocalName != "sst" || written.DocumentElement.NamespaceURI != Main ||
            afterItems.Length == 0 ||
            !int.TryParse(written.DocumentElement.GetAttribute("count"), NumberStyles.None, CultureInfo.InvariantCulture, out var count) ||
            count != newReferences ||
            !int.TryParse(written.DocumentElement.GetAttribute("uniqueCount"), NumberStyles.None, CultureInfo.InvariantCulture, out var unique) ||
            unique != afterItems.Length ||
            written.DocumentElement.Attributes.OfType<XmlAttribute>()
                .Any(attribute => attribute.NamespaceURI != "http://www.w3.org/2000/xmlns/" &&
                    attribute.LocalName is not ("count" or "uniqueCount")) ||
            written.DocumentElement.ChildNodes.OfType<XmlElement>().Any(item => item.LocalName != "si" || item.NamespaceURI != Main))
            issues.Add(new("G5", "UNDECLARED_SHARED_STRING_CHANGE", "new sharedStrings structure"));
    }

    private static void CompareRelationships(ZipArchiveEntry original, ZipArchiveEntry written,
        string workbookPart, string? sharedPart, string? chainPart,
        bool addedShared, bool removedChain, List<GateIssue> issues)
    {
        var before = Load(original);
        var after = Load(written);
        const string sharedType = Office + "/sharedStrings";
        const string chainType = Office + "/calcChain";
        var removed = 0;
        var added = 0;
        foreach (var relationship in before.DocumentElement!.ChildNodes.OfType<XmlElement>().ToArray())
        {
            var id = relationship.GetAttribute("Id");
            var candidate = after.DocumentElement!.ChildNodes.OfType<XmlElement>()
                .FirstOrDefault(element => element.GetAttribute("Id") == id);
            if (candidate is null && removedChain && chainPart is not null &&
                relationship.LocalName == "Relationship" && relationship.NamespaceURI == Relationships &&
                relationship.GetAttribute("Type") == chainType &&
                Resolve(workbookPart, relationship.GetAttribute("Target"))
                    .Equals(chainPart, StringComparison.OrdinalIgnoreCase))
            {
                before.DocumentElement.RemoveChild(relationship);
                removed++;
            }
        }
        foreach (var relationship in after.DocumentElement!.ChildNodes.OfType<XmlElement>().ToArray())
        {
            var id = relationship.GetAttribute("Id");
            var existed = before.DocumentElement!.ChildNodes.OfType<XmlElement>()
                .Any(element => element.GetAttribute("Id") == id);
            if (!existed && addedShared && sharedPart is not null &&
                relationship.LocalName == "Relationship" && relationship.NamespaceURI == Relationships &&
                relationship.GetAttribute("Type") == sharedType && id.Length > 0 &&
                relationship.GetAttribute("TargetMode").Length == 0 &&
                relationship.Attributes.OfType<XmlAttribute>().Count(attribute =>
                    attribute.NamespaceURI != "http://www.w3.org/2000/xmlns/") == 3 &&
                Resolve(workbookPart, relationship.GetAttribute("Target"))
                    .Equals(sharedPart, StringComparison.OrdinalIgnoreCase))
            {
                after.DocumentElement.RemoveChild(relationship);
                added++;
            }
        }
        if (removed != (removedChain ? 1 : 0) || added != (addedShared ? 1 : 0) ||
            !Equivalent(before.DocumentElement, after.DocumentElement))
            issues.Add(new("G5", "UNDECLARED_RELATIONSHIP_CHANGE", original.FullName));
    }

    private static void CompareContentTypes(ZipArchiveEntry original, ZipArchiveEntry written,
        string? sharedPart, string? calcPart, bool addedShared, bool removedChain, List<GateIssue> issues)
    {
        var before = Load(original);
        var after = Load(written);
        var removed = 0;
        var added = 0;
        if (removedChain && calcPart is not null)
            foreach (var element in before.DocumentElement!.ChildNodes.OfType<XmlElement>()
                .Where(element => element.LocalName == "Override" && element.NamespaceURI == Types &&
                    element.GetAttribute("PartName").TrimStart('/').Equals(calcPart, StringComparison.OrdinalIgnoreCase) &&
                    element.GetAttribute("ContentType") ==
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml").ToArray())
            {
                before.DocumentElement.RemoveChild(element);
                removed++;
            }
        if (addedShared && sharedPart is not null)
            foreach (var element in after.DocumentElement!.ChildNodes.OfType<XmlElement>()
                .Where(element => element.LocalName == "Override" && element.NamespaceURI == Types &&
                    element.GetAttribute("PartName") == "/" + sharedPart &&
                    element.GetAttribute("ContentType") ==
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" &&
                    element.Attributes.OfType<XmlAttribute>().Count(attribute =>
                        attribute.NamespaceURI != "http://www.w3.org/2000/xmlns/") == 2).ToArray())
            {
                after.DocumentElement.RemoveChild(element);
                added++;
            }
        if (removed != (removedChain ? 1 : 0) || added != (addedShared ? 1 : 0) ||
            !Equivalent(before.DocumentElement, after.DocumentElement))
            issues.Add(new("G5", "UNDECLARED_CONTENT_TYPE_CHANGE", "[Content_Types].xml"));
    }
}
