using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocLoupe.Excel.Model;
using DocLoupe.Excel.Package;

namespace DocLoupe.Excel.Engine;

public sealed record SetValueOp(string Sheet, string Address, string Kind, string? Value, string RichPolicy = "reject", bool AsText = false, string Operation = "set_value", bool RemoveCell = false, bool KeepCache = false);
public sealed record ExpectedCell(string Sheet, string Address, string Kind, string? Value, bool AllowMissing = false, bool RequireMissing = false, bool KeepCache = false);
public sealed record ApplyResult(IReadOnlyList<ExpectedCell> Intent, IReadOnlyList<ByteEdit> Edits, IReadOnlyList<string> ChangedParts);

public static class SetValueEngine
{
    private static readonly UTF8Encoding Utf8 = new(false, true);
    private const string SharedType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    private const string ChainType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
    private const string SharedContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

    public static ApplyResult Apply(PackageStore store, IReadOnlyList<SetValueOp> operations)
    {
        if (operations.Count == 0) throw new ArgumentException("At least one cell operation is required");
        if (operations.Any(operation => operation.KeepCache && (operation.Operation != "set_formula" || operation.Kind != "formula") ||
            operation.RemoveCell && operation.Operation != "clear" ||
            operation.Operation == "clear" &&
            (operation.Kind != "blank" || operation.Value is not null || operation.RichPolicy != "reject" || operation.AsText)))
            throw new NotSupportedException("Only value-only clear is supported");
        var planned = new Dictionary<string, (byte[] Content, IReadOnlyList<ByteEdit> Edits)>(StringComparer.OrdinalIgnoreCase);
        var intent = new List<ExpectedCell>();
        var strings = new SharedStrings(store);
        var overwrittenFormula = false;
        foreach (var group in operations.GroupBy(operation => store.SheetPart(operation.Sheet), StringComparer.OrdinalIgnoreCase))
        {
            var part = group.Key;
            var original = store.Read(part);
            if (original.Length > 32 * 1024 * 1024) throw new NotSupportedException("Worksheet exceeds P2a DOM limit");
            var lexical = new LexicalPart(original);
            var root = lexical.Document.DocumentElement!;
            if (root.LocalName != "worksheet" || root.NamespaceURI != PackageStore.Main)
                throw new InvalidDataException("Invalid worksheet root");
            var sheetData = Direct(root, "sheetData");
            var rows = sheetData is null ? [] : sheetData.ChildNodes.OfType<XmlElement>().Where(element => element.LocalName == "row" && element.NamespaceURI == PackageStore.Main).ToArray();
            ValidateRows(rows);
            var groups = group.GroupBy(operation => CellAddress.Parse(operation.Address)).OrderBy(group => group.Key.Row).ThenBy(group => group.Key.Column).ToArray();
            if (groups.Any(addressGroup => addressGroup.Count() != 1)) throw new InvalidDataException("Duplicate target cell in batch");
            var createdRows = new Dictionary<int, List<(CellAddress Address, string Xml)>>();
            var newCells = new Dictionary<XmlElement, List<(CellAddress Address, string Xml)>>(ReferenceEqualityComparer.Instance);
            foreach (var addressGroup in groups)
            {
                var operation = addressGroup.Single();
                var address = addressGroup.Key;
                ValidateMergedTarget(root, address);
                if (CellAddress.SheetName(operation.Address) is { } qualifiedSheet && qualifiedSheet != operation.Sheet)
                    throw new InvalidDataException("Qualified target does not match sheet");
                var row = rows.SingleOrDefault(candidate => int.Parse(candidate.GetAttribute("r"), CultureInfo.InvariantCulture) == address.Row);
                var cells = row is null ? [] : row.ChildNodes.OfType<XmlElement>().Where(element => element.LocalName == "c" && element.NamespaceURI == PackageStore.Main).ToArray();
                var cell = cells.SingleOrDefault(candidate => candidate.GetAttribute("r") == address.ToString());
                if (cell is not null)
                {
                    var rawStart = lexical.StartTag(cell);
                    if (cell.ChildNodes.OfType<XmlElement>().Any(child => child.NamespaceURI != PackageStore.Main || child.LocalName is not ("f" or "v" or "is")))
                        throw new NotSupportedException("Unsupported existing cell child; refusing to discard it");
                    var priorFormula = Direct(cell, "f");
                    if (operation.KeepCache && priorFormula is null)
                        throw new NotSupportedException("cache: keep requires an existing formula");
                    if (priorFormula is not null && priorFormula.GetAttribute("t") is "shared" or "array" or "dataTable")
                        throw new NotSupportedException("Formula group edits require P3 reference handling");
                    var match = Regex.Match(rawStart, "(?:^|\\s)r\\s*=\\s*(['\"])(?<reference>.*?)\\1", RegexOptions.CultureInvariant);
                    if (!match.Success || match.Groups["reference"].Value != address.ToString())
                        throw new InvalidDataException("Cell r disagrees with lexical source");
                    overwrittenFormula |= Direct(cell, "f") is not null && operation.Kind != "formula";
                    if (operation.RichPolicy is not ("reject" or "replace")) throw new FormatException("Unknown rich_policy");
                    var existingShared = cell.GetAttribute("t") == "s";
                    var index = existingShared ? int.Parse(Direct(cell, "v")?.InnerText ?? throw new InvalidDataException("Missing shared-string index"), CultureInfo.InvariantCulture) : -1;
                    var inline = Direct(cell, "is");
                    var rich = existingShared && strings.IsRich(index) || inline is not null &&
                        inline.ChildNodes.OfType<XmlElement>().Any(child => child.LocalName is "r" or "rPh" or "phoneticPr");
                    if (rich && operation.RichPolicy != "replace" && operation.Operation != "clear")
                        throw new InvalidDataException("RICH_CONTENT_REQUIRES_REPLACE");
                    if (operation.RemoveCell)
                    {
                        if (existingShared) strings.RemoveReference();
                        lexical.Replace(cell, "");
                    }
                    else if (operation.Operation != "clear" || cell.HasAttribute("t") || cell.ChildNodes.OfType<XmlElement>().Any())
                    {
                        if (existingShared && operation.Kind is not "text") strings.RemoveReference();
                        var markup = MakeCell(lexical.Document, cell, operation, address, strings);
                        lexical.Replace(cell, markup);
                    }
                }
                else if (operation.Operation != "clear")
                {
                    if (operation.KeepCache) throw new NotSupportedException("cache: keep requires an existing formula");
                    var markup = MakeCell(lexical.Document, null, operation, address, strings);
                    if (row is null)
                    {
                        if (!createdRows.TryGetValue(address.Row, out var added)) createdRows[address.Row] = added = [];
                        added.Add((address, markup));
                    }
                    else
                    {
                        if (!newCells.TryGetValue(row, out var added)) newCells[row] = added = [];
                        added.Add((address, markup));
                    }
                }
                intent.Add(new ExpectedCell(operation.Sheet, address.ToString(), operation.Kind, operation.Value,
                    operation.Operation == "clear", operation.RemoveCell, operation.KeepCache));
            }
            foreach (var (row, added) in newCells)
                InsertCells(lexical, row, added);
            if (createdRows.Count > 0)
                InsertRows(lexical, root, sheetData, rows, createdRows);
            var result = lexical.Finish(part);
            if (result.Edits.Count > 0) planned.Add(part, result);
        }
        if (planned.Count > 0)
        {
            strings.Finish(planned);
            UpdateWorkbook(store, planned, overwrittenFormula);
        }
        var edits = planned.Values.SelectMany(value => value.Edits).ToList();
        var deleted = new List<string>();
        if (overwrittenFormula) RemoveCalculationChain(store, planned, edits, deleted);
        foreach (var (part, value) in planned) store.Set(part, value.Content);
        foreach (var part in deleted) store.Delete(part);
        return new ApplyResult(intent, edits, store.ChangedParts.ToArray());
    }

    private static XmlElement? Direct(XmlElement element, string name) => element.ChildNodes.OfType<XmlElement>()
        .FirstOrDefault(child => child.LocalName == name && child.NamespaceURI == PackageStore.Main);

    private static void ValidateMergedTarget(XmlElement root, CellAddress address)
    {
        var merges = Direct(root, "mergeCells");
        if (merges is null) return;
        foreach (var merge in merges.ChildNodes.OfType<XmlElement>()
            .Where(element => element.LocalName == "mergeCell" && element.NamespaceURI == PackageStore.Main))
        {
            var reference = merge.GetAttribute("ref");
            var ends = reference.Split(':');
            if (ends.Length != 2) throw new InvalidDataException($"Invalid merge range: {reference}");
            var origin = CellAddress.Parse(ends[0]);
            var end = CellAddress.Parse(ends[1]);
            if (end.Row < origin.Row || end.Column < origin.Column)
                throw new InvalidDataException($"Invalid merge range: {reference}");
            if (address.Row >= origin.Row && address.Row <= end.Row
                && address.Column >= origin.Column && address.Column <= end.Column && address != origin)
                throw new InvalidDataException($"MERGED_NON_ORIGIN: {address} belongs to {reference}; edit {origin}");
        }
    }

    private static void ValidateRows(XmlElement[] rows)
    {
        var previousRow = 0;
        foreach (var row in rows)
        {
            if (!int.TryParse(row.GetAttribute("r"), out var rowNumber) || rowNumber <= previousRow)
                throw new InvalidDataException("Unsorted or malformed worksheet rows");
            previousRow = rowNumber;
            if (row.ChildNodes.OfType<XmlElement>().Any(child =>
                child.NamespaceURI != PackageStore.Main || Array.IndexOf(SchemaElementOrder.Row, child.LocalName) < 0))
                throw new NotSupportedException("Unmodeled row child");
            var previousColumn = 0;
            foreach (var cell in row.ChildNodes.OfType<XmlElement>().Where(element => element.LocalName == "c" && element.NamespaceURI == PackageStore.Main))
            {
                var address = CellAddress.Parse(cell.GetAttribute("r"));
                if (address.Row != rowNumber || address.Column <= previousColumn || address.ToString() != cell.GetAttribute("r"))
                    throw new InvalidDataException("Unsorted or mismatched cell references");
                previousColumn = address.Column;
            }
        }
    }

    private static string MakeCell(XmlDocument document, XmlElement? original, SetValueOp operation, CellAddress address, SharedStrings strings)
    {
        var cell = original ?? document.CreateElement(document.DocumentElement!.Prefix, "c", PackageStore.Main);
        cell.SetAttribute("r", address.ToString());
        var existing = original?.GetAttribute("t");
        var oldCache = operation.KeepCache ? Direct(original!, "v")?.CloneNode(true) : null;
        var kind = operation.Kind;
        var value = operation.Value;
        if (kind == "text" && value?.StartsWith('=') == true && !operation.AsText)
            throw new InvalidDataException("AMBIGUOUS_FORMULA_TEXT: use explicit kind formula or inline");
        if (kind is not ("text" or "inline" or "number" or "boolean" or "formula" or "blank" or "error"))
            throw new NotSupportedException($"Unsupported set_value kind: {kind}");
        if (kind == "number" && (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number)))
            throw new FormatException("Invalid finite number");
        if (kind == "boolean" && value is not ("true" or "false")) throw new FormatException("Boolean must be true or false");
        if (kind == "error" && !CellError.IsSupported(value)) throw new FormatException("Unsupported Excel error token");
        while (cell.FirstChild is { } child) cell.RemoveChild(child);
        if (kind == "blank") cell.RemoveAttribute("t");
        else if (kind == "formula")
        {
            if (operation.KeepCache)
            {
                if (existing is not ("" or "n" or "b" or "e" or "str"))
                    throw new NotSupportedException("Unsupported formula cache type");
            }
            else cell.RemoveAttribute("t");
            var formulaText = value?.TrimStart('=');
            if (string.IsNullOrWhiteSpace(formulaText)) throw new FormatException("Formula is required");
            var formula = document.CreateElement(cell.Prefix, "f", PackageStore.Main);
            formula.InnerText = formulaText;
            cell.AppendChild(formula);
            if (oldCache is not null) cell.AppendChild(oldCache);
        }
        else if (kind == "inline" || kind == "text" && existing == "inlineStr")
        {
            cell.SetAttribute("t", "inlineStr");
            var inline = document.CreateElement(cell.Prefix, "is", PackageStore.Main);
            var text = document.CreateElement(cell.Prefix, "t", PackageStore.Main);
            text.InnerText = value ?? "";
            if (value?.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
                text.SetAttribute("xml:space", "preserve");
            inline.AppendChild(text);
            cell.AppendChild(inline);
        }
        else
        {
            cell.SetAttribute("t", kind switch { "text" => "s", "boolean" => "b", "error" => "e", _ => "n" });
            var scalar = document.CreateElement(cell.Prefix, "v", PackageStore.Main);
            scalar.InnerText = kind switch { "text" => strings.Index(value ?? "", existing != "s").ToString(CultureInfo.InvariantCulture), "boolean" => value == "true" ? "1" : "0", _ => value! };
            cell.AppendChild(scalar);
        }
        return cell.OuterXml;
    }

    private static void InsertCells(LexicalPart lexical, XmlElement row, List<(CellAddress Address, string Xml)> added)
    {
        var existing = row.ChildNodes.OfType<XmlElement>().Where(element => element.LocalName == "c" && element.NamespaceURI == PackageStore.Main).ToArray();
        foreach (var group in added.GroupBy(item => existing.FirstOrDefault(cell => CellAddress.Parse(cell.GetAttribute("r")).Column > item.Address.Column)))
        {
            var xml = string.Concat(group.OrderBy(item => item.Address.Column).Select(item => item.Xml));
            if (group.Key is not null) lexical.InsertBefore(group.Key, xml);
            else if (row.ChildNodes.OfType<XmlElement>().FirstOrDefault(child =>
                child.LocalName == "extLst" && child.NamespaceURI == PackageStore.Main) is { } extension)
                lexical.InsertBefore(extension, xml);
            else
            {
                var start = lexical.StartTag(row);
                if (start.TrimEnd().EndsWith("/>")) lexical.Replace(row, start[..start.LastIndexOf("/>", StringComparison.Ordinal)] + ">" + xml + $"</{row.Name}>");
                else lexical.AppendChild(row, xml);
            }
        }
    }

    private static void InsertRows(LexicalPart lexical, XmlElement root, XmlElement? sheetData, XmlElement[] rows,
        Dictionary<int, List<(CellAddress Address, string Xml)>> created)
    {
        var rowPrefix = rows.FirstOrDefault()?.Prefix ?? sheetData?.Prefix ?? root.Prefix;
        foreach (var group in created.OrderBy(pair => pair.Key)
            .GroupBy(pair => rows.FirstOrDefault(row => int.Parse(row.GetAttribute("r"), CultureInfo.InvariantCulture) > pair.Key)))
        {
            var xml = string.Concat(group.OrderBy(pair => pair.Key).Select(pair =>
                $"<{(rowPrefix.Length == 0 ? "" : rowPrefix + ":")}row r=\"{pair.Key}\">" +
                string.Concat(pair.Value.OrderBy(cell => cell.Address.Column).Select(cell => cell.Xml)) +
                $"</{(rowPrefix.Length == 0 ? "" : rowPrefix + ":")}row>"));
            if (group.Key is not null) lexical.InsertBefore(group.Key, xml);
            else if (sheetData is not null)
            {
                var start = lexical.StartTag(sheetData);
                if (start.TrimEnd().EndsWith("/>")) lexical.Replace(sheetData,
                    start[..start.LastIndexOf("/>", StringComparison.Ordinal)] + ">" + xml + $"</{sheetData.Name}>");
                else lexical.AppendChild(sheetData, xml);
            }
            else
            {
                var prefix = root.Prefix.Length == 0 ? "" : root.Prefix + ":";
                InsertOrdered(lexical, root, "sheetData", $"<{prefix}sheetData>{xml}</{prefix}sheetData>", SchemaElementOrder.Worksheet);
            }
        }
    }

    private static void InsertOrdered(LexicalPart lexical, XmlElement parent, string name, string markup, string[] order)
    {
        var index = Array.IndexOf(order, name);
        if (index < 0 || parent.ChildNodes.OfType<XmlElement>().Any(child =>
            child.NamespaceURI == PackageStore.Main && Array.IndexOf(order, child.LocalName) < 0))
            throw new NotSupportedException("Element is absent from the generated schema-order table");
        var following = parent.ChildNodes.OfType<XmlElement>().FirstOrDefault(child =>
            child.NamespaceURI == PackageStore.Main && Array.IndexOf(order, child.LocalName) > index);
        if (following is not null) lexical.InsertBefore(following, markup);
        else lexical.AppendChild(parent, markup);
    }

    private static void UpdateWorkbook(PackageStore store, Dictionary<string, (byte[] Content, IReadOnlyList<ByteEdit> Edits)> planned, bool removedFormula)
    {
        var part = store.WorkbookPart;
        var lexical = new LexicalPart(store.Read(part));
        var root = lexical.Document.DocumentElement!;
        var calc = Direct(root, "calcPr");
        if (calc is null)
        {
            var prefix = root.Prefix.Length == 0 ? "" : root.Prefix + ":";
            InsertOrdered(lexical, root, "calcPr", $"<{prefix}calcPr fullCalcOnLoad=\"1\"/>", SchemaElementOrder.Workbook);
        }
        else
        {
            var original = lexical.StartTag(calc);
            var match = Regex.Match(original, "(?<![\\w:])fullCalcOnLoad\\s*=\\s*(['\"]).*?\\1", RegexOptions.CultureInvariant);
            var changed = match.Success ? original[..match.Index] + "fullCalcOnLoad=\"1\"" + original[(match.Index + match.Length)..] :
                original.Insert(original.LastIndexOf(calc.Name, StringComparison.Ordinal) + calc.Name.Length, " fullCalcOnLoad=\"1\"");
            if (changed != original) lexical.ReplaceStartTag(calc, changed);
        }
        var result = lexical.Finish(part);
        if (result.Edits.Count > 0) planned.Add(part, result);
    }

    private static void RemoveCalculationChain(PackageStore store,
        Dictionary<string, (byte[] Content, IReadOnlyList<ByteEdit> Edits)> planned, List<ByteEdit> edits, List<string> deleted)
    {
        var chain = store.ReadRelationships(store.WorkbookPart).FirstOrDefault(relationship => relationship.Type == ChainType);
        if (chain is null) return;
        var part = PackageStore.Resolve(store.WorkbookPart, chain.Target);
        if (!store.Contains(part)) throw new InvalidDataException("Missing calcChain target");
        deleted.Add(part);
        var rels = PackageStore.RelationshipPart(store.WorkbookPart);
        RemoveElement(store, rels, element => element.LocalName == "Relationship" && element.GetAttribute("Id") == chain.Id, planned, edits);
        RemoveElement(store, "[Content_Types].xml", element => element.LocalName == "Override" && element.GetAttribute("PartName").TrimStart('/').Equals(part, StringComparison.OrdinalIgnoreCase), planned, edits);
    }

    private static void RemoveElement(PackageStore store, string part, Func<XmlElement, bool> match,
        Dictionary<string, (byte[] Content, IReadOnlyList<ByteEdit> Edits)> planned, List<ByteEdit> edits)
    {
        var lexical = new LexicalPart(planned.TryGetValue(part, out var value) ? value.Content : store.Read(part));
        var element = lexical.Document.DocumentElement!.ChildNodes.OfType<XmlElement>().SingleOrDefault(match);
        if (element is null) throw new InvalidDataException($"Missing expected entry in {part}");
        lexical.Replace(element, "");
        var result = lexical.Finish(part);
        var rebased = new List<ByteEdit>();
        foreach (var edit in result.Edits)
        {
            var start = OriginalOffset(edit.Start, edits.Where(previous => previous.Part.Equals(part, StringComparison.OrdinalIgnoreCase)));
            var end = OriginalOffset(edit.End, edits.Where(previous => previous.Part.Equals(part, StringComparison.OrdinalIgnoreCase)));
            var original = store.ReadOriginal(part);
            if (end < start || end > original.Length || !original.AsSpan(start, end - start).SequenceEqual(edit.Before))
                throw new InvalidDataException("Removal overlaps another package edit");
            rebased.Add(new ByteEdit(part, start, end, edit.Before, edit.After));
        }
        edits.AddRange(rebased);
        planned[part] = (result.Content, (value.Edits ?? []).Concat(rebased).ToArray());
    }

    private static int OriginalOffset(int writtenOffset, IEnumerable<ByteEdit> previous)
    {
        var shift = 0;
        foreach (var edit in previous.OrderBy(item => item.Start).ThenBy(item => item.End))
        {
            var start = edit.Start + shift;
            var end = start + edit.After.Length;
            if (writtenOffset > start && writtenOffset < end)
                throw new InvalidDataException("Package edits overlap");
            if (writtenOffset < end) break;
            shift += edit.After.Length - (edit.End - edit.Start);
        }
        return writtenOffset - shift;
    }

    private sealed class SharedStrings
    {
        private readonly PackageStore _store;
        private readonly string _part;
        private readonly string? _relationship;
        private readonly List<string?> _values = [];
        private readonly List<string> _added = [];
        private int _references;

        public SharedStrings(PackageStore store)
        {
            _store = store;
            var relationship = store.ReadRelationships(store.WorkbookPart).FirstOrDefault(item => item.Type == SharedType);
            _relationship = relationship?.Id;
            var workbookDirectory = store.WorkbookPart.Contains('/') ? store.WorkbookPart[..(store.WorkbookPart.LastIndexOf('/') + 1)] : "";
            _part = relationship is null ? workbookDirectory + "sharedStrings.xml" : PackageStore.Resolve(store.WorkbookPart, relationship.Target);
            if (relationship is null) return;
            var document = PackageStore.Parse(store.Read(_part));
            foreach (var item in document.DocumentElement!.ChildNodes.OfType<XmlElement>().Where(item => item.LocalName == "si" && item.NamespaceURI == PackageStore.Main))
                _values.Add(item.ChildNodes.OfType<XmlElement>().Any(child => child.LocalName is "r" or "rPh" or "phoneticPr") ? null :
                    item.GetElementsByTagName("t", PackageStore.Main).OfType<XmlElement>().FirstOrDefault()?.InnerText);
        }

        public bool IsRich(int index) => index < 0 || index >= _values.Count
            ? throw new InvalidDataException("Invalid shared-string index") : _values[index] is null;

        public void RemoveReference() => _references--;

        public int Index(string text, bool newReference)
        {
            if (newReference) _references++;
            var index = _values.IndexOf(text);
            if (index >= 0) return index;
            _values.Add(text);
            _added.Add(text);
            return _values.Count - 1;
        }

        public void Finish(Dictionary<string, (byte[] Content, IReadOnlyList<ByteEdit> Edits)> planned)
        {
            if (_references == 0 && _added.Count == 0) return;
            if (_relationship is null)
            {
                if (_store.Contains(_part)) throw new InvalidDataException("Unrelated sharedStrings part already exists");
                var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"" + PackageStore.Main + "\" count=\"" + _references + "\" uniqueCount=\"" + _added.Count + "\">" +
                    string.Concat(_added.Select(text => WriteItem(text))) + "</sst>";
                planned.Add(_part, (Utf8.GetBytes(xml), []));
                var relationships = PackageStore.RelationshipPart(_store.WorkbookPart);
                var rel = new LexicalPart(_store.Read(relationships));
                var root = rel.Document.DocumentElement!;
                var used = _store.ReadRelationships(_store.WorkbookPart).Select(item => item.Id).ToHashSet();
                var id = 1;
                while (used.Contains("rId" + id)) id++;
                rel.AppendChild(root, $"<Relationship Id=\"rId{id}\" Type=\"{SharedType}\" Target=\"sharedStrings.xml\" xmlns=\"{PackageStore.PackageRelationships}\"/>");
                planned.Add(relationships, rel.Finish(relationships));
                var types = new LexicalPart(_store.Read("[Content_Types].xml"));
                types.AppendChild(types.Document.DocumentElement!, $"<Override PartName=\"/{_part}\" ContentType=\"{SharedContentType}\" xmlns=\"{PackageStore.ContentTypes}\"/>");
                planned.Add("[Content_Types].xml", types.Finish("[Content_Types].xml"));
            }
            else
            {
                var lexical = new LexicalPart(_store.Read(_part));
                var root = lexical.Document.DocumentElement!;
                var raw = lexical.StartTag(root);
                foreach (var (attribute, delta) in new[] { ("count", _references), ("uniqueCount", _added.Count) })
                {
                    var match = Regex.Match(raw, $"(?<![\\w:]){attribute}\\s*=\\s*(['\"])(?<number>\\d+)\\1");
                    if (!match.Success) continue;
                    var value = checked(int.Parse(match.Groups["number"].Value, CultureInfo.InvariantCulture) + delta);
                    raw = raw[..match.Groups["number"].Index] + value.ToString(CultureInfo.InvariantCulture) + raw[(match.Groups["number"].Index + match.Groups["number"].Length)..];
                }
                lexical.ReplaceStartTag(root, raw);
                if (_added.Count > 0) lexical.AppendChild(root, string.Concat(_added.Select(text => WriteItem(text, root.Prefix))));
                planned.Add(_part, lexical.Finish(_part));
            }
        }

        private static string WriteItem(string text, string prefix = "")
        {
            var tag = prefix.Length == 0 ? "" : prefix + ":";
            var document = new XmlDocument();
            var item = document.CreateElement(prefix, "si", PackageStore.Main);
            var value = document.CreateElement(prefix, "t", PackageStore.Main);
            value.InnerText = text;
            if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
                value.SetAttribute("xml:space", "preserve");
            item.AppendChild(value);
            return item.OuterXml.Replace($"<{tag}si xmlns=\"{PackageStore.Main}\"", $"<{tag}si", StringComparison.Ordinal);
        }
    }
}
