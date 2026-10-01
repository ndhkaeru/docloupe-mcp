using System.IO.Compression;
using System.Text;
using System.Xml;
using DocLoupe.Excel.Verify;
using DocumentFormat.OpenXml.Spreadsheet;

internal sealed record EditResult(int ParsedPartCount, int TypedCellCount, string? CellReference, bool DeclarationPreserved,
    bool RootPrefixPreserved, bool RootNamespacesPreserved, bool UntouchedPartsPreserved,
    bool UntouchedSheetBytesPreserved, bool EditedCellPresent, MarkupResult G3);

internal static class RawWorksheetEditor
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Marker = "S2b-edited";

    public static EditResult Edit(string source, string output)
    {
        using var input = ZipFile.OpenRead(source);
        var xmlParts = input.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToArray();
        foreach (var entry in xmlParts)
            LoadPart(entry);

        var sheetEntry = input.Entries.First(entry => entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal)
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
        var sourceBytes = Read(sheetEntry);
        var sourceDocument = LoadPart(sheetEntry);
        var worksheet = new Worksheet();
        foreach (XmlAttribute attribute in sourceDocument.DocumentElement!.Attributes)
        {
            if (attribute.Prefix == "xmlns")
                worksheet.AddNamespaceDeclaration(attribute.LocalName, attribute.Value);
        }
        worksheet.InnerXml = sourceDocument.DocumentElement.InnerXml;
        var typedCell = worksheet.Descendants<Cell>().First();
        var cellReference = typedCell.CellReference?.Value;
        var xmlCell = sourceDocument.GetElementsByTagName("c", MainNamespace).OfType<XmlElement>().First();
        if (cellReference != xmlCell.GetAttribute("r"))
            throw new InvalidOperationException("Detached typed DOM and raw XML disagree on the edited cell");

        typedCell.DataType = CellValues.InlineString;
        typedCell.CellValue = null;
        typedCell.InlineString = new InlineString(new Text(Marker));
        var editedText = typedCell.InlineString.Text?.Text
            ?? throw new InvalidOperationException("Typed DOM edit did not produce text");
        xmlCell.SetAttribute("t", "inlineStr");
        while (xmlCell.FirstChild is { } child)
            xmlCell.RemoveChild(child);
        var inline = sourceDocument.CreateElement(xmlCell.Prefix, "is", MainNamespace);
        var text = sourceDocument.CreateElement(xmlCell.Prefix, "t", MainNamespace);
        text.InnerText = editedText;
        inline.AppendChild(text);
        xmlCell.AppendChild(inline);
        var serialized = Serialize(sourceBytes, sourceDocument);
        var outputBytes = ReplaceCellOnly(sourceBytes, serialized, xmlCell.Name);

        using (var destination = new ZipArchive(File.Create(output), ZipArchiveMode.Create))
        {
            foreach (var entry in input.Entries)
            {
                var target = destination.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                target.LastWriteTime = entry.LastWriteTime;
                using var targetStream = target.Open();
                if (entry.FullName == sheetEntry.FullName)
                    targetStream.Write(outputBytes);
                else
                    using (var sourceStream = entry.Open()) sourceStream.CopyTo(targetStream);
            }
        }

        using var written = ZipFile.OpenRead(output);
        var writtenBytes = Read(written.GetEntry(sheetEntry.FullName)!);
        var untouched = input.Entries.Where(entry => entry.FullName != sheetEntry.FullName)
            .All(entry => written.GetEntry(entry.FullName) is { } after && Read(entry).AsSpan().SequenceEqual(Read(after)));
        var newDocument = LoadPart(written.GetEntry(sheetEntry.FullName)!);
        var newCell = newDocument.GetElementsByTagName("c", MainNamespace).OfType<XmlElement>().First();
        var originalText = new UTF8Encoding(false, true).GetString(sourceBytes);
        var writtenText = new UTF8Encoding(false, true).GetString(writtenBytes);
        var originalSpan = FindFirstElement(originalText, xmlCell.Name);
        var writtenSpan = FindFirstElement(writtenText, xmlCell.Name);
        var outsideCell = originalText.AsSpan(0, originalSpan.Start).SequenceEqual(writtenText.AsSpan(0, writtenSpan.Start))
            && originalText.AsSpan(originalSpan.End).SequenceEqual(writtenText.AsSpan(writtenSpan.End));
        using var originalStream = new MemoryStream(sourceBytes);
        using var writtenStream = new MemoryStream(writtenBytes);
        var g3 = MarkupCompatibilityVerifier.Check(writtenStream, originalStream, cellReference);
        return new EditResult(xmlParts.Length, worksheet.Descendants<Cell>().Count(), cellReference,
            Declaration(sourceBytes) == Declaration(writtenBytes),
            sourceDocument.DocumentElement.Prefix == newDocument.DocumentElement!.Prefix,
            Declarations(sourceDocument.DocumentElement).SequenceEqual(Declarations(newDocument.DocumentElement)),
            untouched, outsideCell, newCell.InnerText == Marker, g3);
    }

    internal static XmlDocument LoadPart(ZipArchiveEntry entry)
    {
        var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        document.Load(reader);
        return document;
    }

    private static byte[] Serialize(byte[] source, XmlDocument document)
    {
        var original = new UTF8Encoding(false, true).GetString(source);
        if (original.AsSpan(0, Math.Min(64, original.Length)).Contains('\0'))
            throw new NotSupportedException("This spike only writes UTF-8 worksheet XML");
        var bom = original.Length > 0 && original[0] == '\uFEFF' ? "\uFEFF" : "";
        var xml = bom + Declaration(source) + string.Concat(document.ChildNodes.OfType<XmlNode>()
            .Where(node => node.NodeType != XmlNodeType.XmlDeclaration).Select(node => node.OuterXml));
        return new UTF8Encoding(false, true).GetBytes(xml);
    }

    private static byte[] ReplaceCellOnly(byte[] originalBytes, byte[] serializedBytes, string qualifiedName)
    {
        var encoding = new UTF8Encoding(false, true);
        var original = encoding.GetString(originalBytes);
        var serialized = encoding.GetString(serializedBytes);
        var originalSpan = FindFirstElement(original, qualifiedName);
        var serializedSpan = FindFirstElement(serialized, qualifiedName);
        var output = string.Concat(original.AsSpan(0, originalSpan.Start),
            serialized.AsSpan(serializedSpan.Start, serializedSpan.End - serializedSpan.Start),
            original.AsSpan(originalSpan.End));
        return encoding.GetBytes(output);
    }

    private static (int Start, int End) FindFirstElement(string xml, string qualifiedName)
    {
        var offset = 0;
        var start = -1;
        var depth = 0;
        while (offset < xml.Length)
        {
            var opening = xml.IndexOf('<', offset);
            if (opening < 0) break;
            if (xml.AsSpan(opening).StartsWith("<!--", StringComparison.Ordinal))
            {
                offset = SkipMarkup(xml, opening, "-->");
                continue;
            }
            if (xml.AsSpan(opening).StartsWith("<![CDATA[", StringComparison.Ordinal))
            {
                offset = SkipMarkup(xml, opening, "]]>");
                continue;
            }
            if (xml.AsSpan(opening).StartsWith("<?", StringComparison.Ordinal))
            {
                offset = SkipMarkup(xml, opening, "?>");
                continue;
            }
            if (opening + 1 >= xml.Length || xml[opening + 1] == '!')
                throw new XmlException("Unexpected markup in worksheet XML");

            var closing = xml[opening + 1] == '/';
            var nameStart = opening + (closing ? 2 : 1);
            var nameEnd = nameStart;
            while (nameEnd < xml.Length && !char.IsWhiteSpace(xml[nameEnd]) && xml[nameEnd] is not '/' and not '>')
                nameEnd++;
            var name = xml[nameStart..nameEnd];
            var end = FindTagEnd(xml, nameEnd);
            if (closing && start >= 0)
            {
                depth--;
                if (depth == 0) return (start, end);
            }
            else if (!closing)
            {
                var last = end - 2;
                while (last > nameEnd && char.IsWhiteSpace(xml[last])) last--;
                var empty = xml[last] == '/';
                if (start >= 0 && !empty) depth++;
                else if (start < 0 && name == qualifiedName)
                {
                    if (empty) return (opening, end);
                    start = opening;
                    depth = 1;
                }
            }
            offset = end;
        }
        throw new XmlException($"Could not locate complete {qualifiedName} element");
    }

    private static int FindTagEnd(string xml, int offset)
    {
        var quote = '\0';
        for (var index = offset; index < xml.Length; index++)
        {
            var character = xml[index];
            if (quote != '\0')
            {
                if (character == quote) quote = '\0';
            }
            else if (character is '\'' or '"') quote = character;
            else if (character == '>') return index + 1;
        }
        throw new XmlException("Unterminated XML element");
    }

    private static int SkipMarkup(string xml, int offset, string terminator)
    {
        var end = xml.IndexOf(terminator, offset, StringComparison.Ordinal);
        if (end < 0) throw new XmlException("Unterminated XML markup");
        return end + terminator.Length;
    }

    internal static string Declaration(byte[] bytes)
    {
        var xml = new UTF8Encoding(false, true).GetString(bytes);
        var start = xml.StartsWith('\uFEFF') ? 1 : 0;
        while (start < xml.Length && char.IsWhiteSpace(xml[start])) start++;
        if (!xml.AsSpan(start).StartsWith("<?xml", StringComparison.OrdinalIgnoreCase)
            || start + 5 >= xml.Length || !char.IsWhiteSpace(xml[start + 5])) return "";
        var end = xml.IndexOf("?>", start, StringComparison.Ordinal);
        if (end < 0) throw new XmlException("Unterminated XML declaration");
        return xml[start..(end + 2)];
    }

    private static IEnumerable<string> Declarations(XmlElement element) => element.Attributes.OfType<XmlAttribute>()
        .Where(attribute => attribute.Name == "xmlns" || attribute.Prefix == "xmlns")
        .Select(attribute => $"{attribute.Name}={attribute.Value}").OrderBy(value => value, StringComparer.Ordinal);

    private static byte[] Read(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }
}
