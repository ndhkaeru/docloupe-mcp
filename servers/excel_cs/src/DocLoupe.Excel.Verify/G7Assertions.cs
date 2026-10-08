using System.Globalization;
using System.Text.RegularExpressions;

namespace DocLoupe.Excel.Verify;

public sealed record ValueAssertion(string Sheet, string Address, bool CheckValue, string? Kind, string? Value, string? Formula, bool Unchanged = false, string? Rich = null, bool? FontBold = null, bool? FontItalic = null);

public static class G7Assertions
{
    private static readonly Regex NumericPattern = new(@"\A(?<sign>[+-]?)(?<whole>[0-9]+)(?:\.(?<fraction>[0-9]+))?(?:[eE](?<exponent>[+-]?[0-9]+))?\z",
        RegexOptions.CultureInvariant | RegexOptions.NonBacktracking);

    public static bool IsSupportedNumber(string? value) => CanonicalNumber(value) is not null;

    public static IReadOnlyList<GateIssue> Check(string path, IReadOnlyList<ValueAssertion> assertions, string? source = null)
    {
        var issues = new List<GateIssue>();
        foreach (var group in assertions.GroupBy(item => item.Sheet, StringComparer.Ordinal))
        {
            Dictionary<string, CellRead> cells;
            try
            {
                cells = P2aGates.ReadCells(path, group.Key, group.Select(item => item.Address))
                    .ToDictionary(cell => cell.Address, StringComparer.Ordinal);
            }
            catch (InvalidOperationException)
            {
                issues.AddRange(group.Select(item => new GateIssue("G7", "ASSERT_SHEET_MISSING", $"{group.Key}!{item.Address}")));
                continue;
            }
            Dictionary<string, CellRead>? originals = null;
            IReadOnlyDictionary<string, ExplicitFontStyle>? fontStyle = null;
            if (group.Any(item => item.FontBold is not null || item.FontItalic is not null))
                fontStyle = ReadFontStyle(path, group.Key, cells.Values);
            if (source is not null && group.Any(item => item.Unchanged))
            {
                try
                {
                    originals = P2aGates.ReadCells(source, group.Key, group.Where(item => item.Unchanged)
                        .Select(item => item.Address)).ToDictionary(cell => cell.Address, StringComparer.Ordinal);
                }
                catch (InvalidOperationException)
                {
                    issues.AddRange(group.Where(item => item.Unchanged).Select(item =>
                        new GateIssue("G7", "ASSERT_SOURCE_SHEET_MISSING", $"{group.Key}!{item.Address}")));
                }
            }
            foreach (var assertion in group)
            {
                cells.TryGetValue(assertion.Address, out var actual);
                var target = $"{group.Key}!{assertion.Address}";
                if (assertion.Unchanged)
                {
                    if (source is null) issues.Add(new GateIssue("G7", "ASSERT_SOURCE_REQUIRED", target));
                    else if (originals is not null)
                    {
                        originals.TryGetValue(assertion.Address, out var original);
                        if (original != actual) issues.Add(new GateIssue("G7", "ASSERT_CHANGED", target));
                    }
                    continue;
                }
                issues.AddRange(CheckValueAndFormula(actual, assertion, target));
                if (assertion.Rich is { } rich && !RichTextAssertions.Matches(actual, rich))
                    issues.Add(new GateIssue("G7", "ASSERT_RICH_MISMATCH", target));
                if (assertion.FontBold is not null || assertion.FontItalic is not null)
                {
                    var observedStyle = fontStyle?.GetValueOrDefault(assertion.Address);
                    if (assertion.FontBold is not null && observedStyle?.Bold is null ||
                        assertion.FontItalic is not null && observedStyle?.Italic is null)
                        issues.Add(new GateIssue("G7", "ASSERT_STYLE_UNVERIFIED", target));
                    else if (assertion.FontBold is { } expectedBold && observedStyle?.Bold != expectedBold ||
                             assertion.FontItalic is { } expectedItalic && observedStyle?.Italic != expectedItalic)
                        issues.Add(new GateIssue("G7", "ASSERT_STYLE_MISMATCH", target));
                }
            }
        }
        return issues;
    }

    public static bool ValueMatches(CellRead? actual, ValueAssertion assertion)
    {
        if (actual?.Kind == "formula") return FormulaCacheMatches(actual, assertion);
        if (assertion.Kind == "blank") return actual is null || actual.Kind == "blank";
        if (actual is null) return false;
        if (assertion.Kind == "text")
            return actual.Kind is "text" or "inline" && actual.Value == assertion.Value;
        if (assertion.Kind == "number")
            return actual.Kind == "number" && CanonicalNumber(actual.Value) is { } numeric &&
                CanonicalNumber(assertion.Value) is { } expected && numeric == expected;
        return assertion.Kind == actual.Kind && assertion.Value == actual.Value;
    }

    public static bool Matches(CellRead? actual, ValueAssertion assertion) =>
        !assertion.Unchanged && assertion.FontBold is null && assertion.FontItalic is null && !CheckValueAndFormula(actual, assertion,
            assertion.Sheet + "!" + assertion.Address).Any() &&
        (assertion.Rich is null || RichTextAssertions.Matches(actual, assertion.Rich));

    public static IReadOnlyDictionary<string, bool?> ReadFontBold(string path, string sheetName, IEnumerable<CellRead> cells) =>
        ReadFontStyle(path, sheetName, cells).ToDictionary(item => item.Key, item => item.Value.Bold, StringComparer.Ordinal);

    public static IReadOnlyDictionary<string, ExplicitFontStyle> ReadFontStyle(string path, string sheetName, IEnumerable<CellRead> cells)
    {
        var observed = cells.ToArray();
        try { return P2aGates.ReadExplicitFontStyle(path, sheetName, observed); }
        catch (Exception error) when (error is InvalidDataException or System.Xml.XmlException or
                                      InvalidOperationException or KeyNotFoundException or UriFormatException or ArgumentException)
        {
            return observed.ToDictionary(cell => cell.Address, _ => new ExplicitFontStyle(null, null), StringComparer.Ordinal);
        }
    }

    private static IEnumerable<GateIssue> CheckValueAndFormula(CellRead? actual,
        ValueAssertion assertion, string target)
    {
        if (assertion.Formula is not null && actual?.Formula != assertion.Formula)
            yield return new GateIssue("G7", "ASSERT_FORMULA_MISMATCH", target);
        if (!assertion.CheckValue) yield break;
        if (actual?.Kind == "formula" && actual.CacheType is not ("" or "n" or "b" or "e" or "str"))
            yield return new GateIssue("G7", "ASSERT_CACHE_UNSUPPORTED", target);
        else if (!ValueMatches(actual, assertion))
            yield return new GateIssue("G7", "ASSERT_VALUE_MISMATCH", target);
    }

    private static bool FormulaCacheMatches(CellRead actual, ValueAssertion assertion)
    {
        var raw = actual.CacheRawValue;
        if (assertion.Kind == "blank") return raw is null && actual.CacheType is ("" or "n");
        if (raw is null) return false;
        return assertion.Kind switch
        {
            "number" => actual.CacheType is "" or "n" && CanonicalNumber(raw) is { } numeric &&
                CanonicalNumber(assertion.Value) is { } expected && numeric == expected,
            "text" => actual.CacheType == "str" && raw == assertion.Value,
            "boolean" => actual.CacheType == "b" && raw is ("0" or "1") &&
                (raw == "1" ? "true" : "false") == assertion.Value,
            "error" => actual.CacheType == "e" && DocLoupe.Excel.Model.CellError.IsSupported(raw) &&
                raw == assertion.Value,
            _ => false
        };
    }

    private static (string Digits, long Power)? CanonicalNumber(string? value)
    {
        if (value is null || value.Length > 128) return null;
        var match = NumericPattern.Match(value);
        if (!match.Success) return null;
        var exponentText = match.Groups["exponent"].Value;
        if (exponentText.Length > 0 && !int.TryParse(exponentText, NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture, out _)) return null;
        var exponent = exponentText.Length == 0 ? 0 : int.Parse(exponentText, CultureInfo.InvariantCulture);
        var fraction = match.Groups["fraction"].Value;
        var digits = (match.Groups["whole"].Value + fraction).TrimStart('0');
        if (digits.Length == 0) return ("0", 0);
        var trailing = digits.Length - digits.TrimEnd('0').Length;
        var sign = match.Groups["sign"].Value == "-" ? "-" : "";
        return (sign + digits[..(digits.Length - trailing)], (long)exponent - fraction.Length + trailing);
    }
}
