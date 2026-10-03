using System.Text;
using System.Xml;
using DocLoupe.Excel.Package;

namespace DocLoupe.Excel.Engine;

public sealed record ByteEdit(string Part, int Start, int End, byte[] Before, byte[] After);

internal sealed class LexicalPart
{
    private readonly byte[] _original;
    private readonly string _text;
    private readonly Dictionary<XmlElement, ElementSpan> _locations = new(ReferenceEqualityComparer.Instance);
    private readonly List<(int Start, int End, string Replacement)> _patches = [];
    private static readonly UTF8Encoding Utf8 = new(false, true);

    public LexicalPart(byte[] content)
    {
        _original = content;
        _text = Utf8.GetString(content);
        if (_text.Contains('\0')) throw new NotSupportedException("Only UTF-8 XML parts can be spliced");
        var declarationStart = _text.Length > 0 && _text[0] == '\uFEFF' ? 1 : 0;
        if (_text.AsSpan(declarationStart).StartsWith("<?xml", StringComparison.OrdinalIgnoreCase))
        {
            var declarationEnd = _text.IndexOf("?>", declarationStart, StringComparison.Ordinal);
            if (declarationEnd < 0) throw new InvalidDataException("Unterminated XML declaration");
            var declaration = _text[declarationStart..declarationEnd];
            var encoding = System.Text.RegularExpressions.Regex.Match(declaration,
                "(?<![\\w:])encoding\\s*=\\s*(['\"])(?<value>[^'\"]+)\\1");
            if (encoding.Success && !encoding.Groups["value"].Value.Equals("UTF-8", StringComparison.OrdinalIgnoreCase) &&
                !encoding.Groups["value"].Value.Equals("UTF8", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Only UTF-8 XML parts can be spliced");
        }
        Document = PackageStore.Parse(content);
        Index();
    }

    public XmlDocument Document { get; }

    public void Replace(XmlElement element, string replacement)
    {
        var span = _locations[element];
        Add(span.Start, span.End, replacement);
    }

    public void ReplaceStartTag(XmlElement element, string replacement)
    {
        var span = _locations[element];
        Add(span.Start, span.StartTagEnd, replacement);
    }

    public void InsertBefore(XmlElement element, string markup) => Add(_locations[element].Start, _locations[element].Start, markup);

    public void AppendChild(XmlElement parent, string markup)
    {
        var span = _locations[parent];
        if (span.Empty) throw new InvalidDataException($"Cannot insert into self-closing {parent.Name}");
        Add(span.EndTagStart, span.EndTagStart, markup);
    }

    public string StartTag(XmlElement element)
    {
        var span = _locations[element];
        return _text[span.Start..span.StartTagEnd];
    }

    public (byte[] Content, IReadOnlyList<ByteEdit> Edits) Finish(string part)
    {
        var edits = new List<ByteEdit>();
        var output = _original;
        var prior = _text.Length + 1;
        foreach (var patch in _patches.OrderByDescending(patch => patch.Start).ThenByDescending(patch => patch.End))
        {
            if (patch.End > prior) throw new InvalidDataException("Overlapping edit regions");
            var start = Utf8.GetByteCount(_text.AsSpan(0, patch.Start));
            var end = Utf8.GetByteCount(_text.AsSpan(0, patch.End));
            var replacement = Utf8.GetBytes(patch.Replacement);
            edits.Add(new ByteEdit(part, start, end, _original[start..end], replacement));
            var merged = new byte[output.Length - (end - start) + replacement.Length];
            Buffer.BlockCopy(output, 0, merged, 0, start);
            Buffer.BlockCopy(replacement, 0, merged, start, replacement.Length);
            Buffer.BlockCopy(output, end, merged, start + replacement.Length, output.Length - end);
            output = merged;
            prior = patch.Start;
        }
        return (output, edits);
    }

    private void Add(int start, int end, string replacement)
    {
        if (start < 0 || end < start || end > _text.Length) throw new InvalidDataException("Invalid edit span");
        if (_patches.Any(patch => start < patch.End && end > patch.Start
            || start > patch.Start && start < patch.End
            || start == patch.Start && start == end && patch.Start == patch.End))
            throw new InvalidDataException("Overlapping or ambiguous edits");
        _patches.Add((start, end, replacement));
    }

    private void Index()
    {
        var elements = Descendants(Document.DocumentElement ?? throw new InvalidDataException("Missing XML root")).GetEnumerator();
        var stack = new Stack<(XmlElement Element, int Start, int TagEnd)>();
        for (var offset = 0; offset < _text.Length;)
        {
            var opening = _text.IndexOf('<', offset);
            if (opening < 0) break;
            if (_text.AsSpan(opening).StartsWith("<!--")) { offset = Skip(opening, "-->"); continue; }
            if (_text.AsSpan(opening).StartsWith("<![CDATA[")) { offset = Skip(opening, "]]>"); continue; }
            if (_text.AsSpan(opening).StartsWith("<?")) { offset = Skip(opening, "?>"); continue; }
            if (_text.AsSpan(opening).StartsWith("<!")) throw new InvalidDataException("DTD and declarations are not supported");
            var closing = _text[opening + 1] == '/';
            var nameStart = opening + (closing ? 2 : 1);
            var nameEnd = nameStart;
            while (nameEnd < _text.Length && !char.IsWhiteSpace(_text[nameEnd]) && _text[nameEnd] is not '/' and not '>') nameEnd++;
            var name = _text[nameStart..nameEnd];
            var tagEnd = TagEnd(nameEnd);
            if (closing)
            {
                if (!stack.TryPop(out var frame) || frame.Element.Name != name)
                    throw new InvalidDataException("XML lexical index disagrees with DOM");
                _locations.Add(frame.Element, new ElementSpan(frame.Start, frame.TagEnd, opening, tagEnd, false));
            }
            else
            {
                if (!elements.MoveNext() || elements.Current.Name != name)
                    throw new InvalidDataException("XML lexical index disagrees with DOM");
                var last = tagEnd - 2;
                while (char.IsWhiteSpace(_text[last])) last--;
                if (_text[last] == '/') _locations.Add(elements.Current, new ElementSpan(opening, tagEnd, opening, tagEnd, true));
                else stack.Push((elements.Current, opening, tagEnd));
            }
            offset = tagEnd;
        }
        if (stack.Count != 0 || elements.MoveNext()) throw new InvalidDataException("Incomplete XML lexical index");
    }

    private int Skip(int start, string terminator)
    {
        var end = _text.IndexOf(terminator, start, StringComparison.Ordinal);
        if (end < 0) throw new InvalidDataException("Incomplete XML markup");
        return end + terminator.Length;
    }

    private int TagEnd(int start)
    {
        var quote = '\0';
        for (var offset = start; offset < _text.Length; offset++)
        {
            if (quote != '\0') { if (_text[offset] == quote) quote = '\0'; }
            else if (_text[offset] is '\'' or '"') quote = _text[offset];
            else if (_text[offset] == '>') return offset + 1;
        }
        throw new InvalidDataException("Unterminated XML tag");
    }

    private static IEnumerable<XmlElement> Descendants(XmlElement element)
    {
        yield return element;
        foreach (var child in element.ChildNodes.OfType<XmlElement>())
        foreach (var descendant in Descendants(child)) yield return descendant;
    }

    private sealed record ElementSpan(int Start, int StartTagEnd, int EndTagStart, int End, bool Empty);
}
