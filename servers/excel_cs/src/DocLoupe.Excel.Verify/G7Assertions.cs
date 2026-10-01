using System.Globalization;
using System.Text.RegularExpressions;

namespace DocLoupe.Excel.Verify;

public sealed record ValueAssertion(string Sheet, string Address, bool CheckValue, string? Kind, string? Value, string? Formula);

public static class G7Assertions
{
    private static readonly Regex NumericPattern = new(@"\A(?<sign>[+-]?)(?<whole>[0-9]+)(?:\.(?<fraction>[0-9]+))?(?:[eE](?<exponent>[+-]?[0-9]+))?\z",
        RegexOptions.CultureInvariant | RegexOptions.NonBacktracking);

    public static bool IsSupportedNumber(string? value) => CanonicalNumber(value) is not null;

    public static IReadOnlyList<GateIssue> Check(string path, IReadOnlyList<ValueAssertion> assertions)
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
            foreach (var assertion in group)
            {
                cells.TryGetValue(assertion.Address, out var actual);
                var target = $"{group.Key}!{assertion.Address}";
                if (assertion.Formula is not null && actual?.Formula != assertion.Formula)
                    issues.Add(new GateIssue("G7", "ASSERT_FORMULA_MISMATCH", target));
                if (assertion.CheckValue && actual?.Kind == "formula")
                    issues.Add(new GateIssue("G7", "ASSERT_CACHE_UNSUPPORTED", target));
                else if (assertion.CheckValue && !ValueMatches(actual, assertion))
                    issues.Add(new GateIssue("G7", "ASSERT_VALUE_MISMATCH", target));
            }
        }
        return issues;
    }

    private static bool ValueMatches(CellRead? actual, ValueAssertion assertion)
    {
        if (assertion.Kind == "blank") return actual is null || actual.Kind == "blank";
        if (actual is null) return false;
        if (assertion.Kind == "text")
            return actual.Kind is "text" or "inline" && actual.Value == assertion.Value;
        if (assertion.Kind == "number")
            return actual.Kind == "number" && CanonicalNumber(actual.Value) is { } numeric &&
                CanonicalNumber(assertion.Value) is { } expected && numeric == expected;
        return assertion.Kind == actual.Kind && assertion.Value == actual.Value;
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
