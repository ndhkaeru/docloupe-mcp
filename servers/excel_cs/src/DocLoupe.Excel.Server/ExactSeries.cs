using System.Globalization;
using System.Numerics;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace DocLoupe.Excel.Server;

internal sealed class ExactSeries
{
    private static readonly Regex NumberPattern = new(
        @"\A(?<sign>-?)(?<whole>0|[1-9][0-9]*)(?:\.(?<fraction>[0-9]+))?(?:[eE](?<exponent>[+-]?[0-9]+))?\z",
        RegexOptions.CultureInvariant | RegexOptions.NonBacktracking);
    private readonly Term _start;
    private readonly Term _step;

    private ExactSeries(Term start, Term step)
    {
        _start = start;
        _step = step;
    }

    public static ExactSeries Parse(JsonElement series)
    {
        if (series.ValueKind != JsonValueKind.Object) throw new NotSupportedException("fill series must be an object");
        var fields = series.EnumerateObject().ToArray();
        if (fields.Length != 2 || fields.Any(field => field.Name is not ("start" or "step")) ||
            !series.TryGetProperty("start", out var start) || !series.TryGetProperty("step", out var step))
            throw new NotSupportedException("fill series requires numeric start and step");
        return new ExactSeries(ParseTerm(start), ParseTerm(step));
    }

    public string At(long index)
    {
        var power = Math.Min(_start.Power, _step.Power);
        var coefficient = _start.Coefficient * BigInteger.Pow(10, _start.Power - power) +
                          _step.Coefficient * BigInteger.Pow(10, _step.Power - power) * index;
        if (coefficient.IsZero) return "0";
        var absolute = BigInteger.Abs(coefficient);
        if (absolute >= BigInteger.Pow(10, 15 - power))
            throw new ArgumentOutOfRangeException(nameof(index), "Series exceeds 15 decimal places before the decimal point");
        while (absolute % 10 == 0)
        {
            absolute /= 10;
            power++;
        }
        var digits = absolute.ToString(CultureInfo.InvariantCulture);
        if (digits.Length > 15)
            throw new ArgumentOutOfRangeException(nameof(index), "Series exceeds 15 significant digits");
        var sign = coefficient.Sign < 0 ? "-" : "";
        if (power >= 0) return sign + digits + new string('0', power);
        var integralLength = digits.Length + power;
        return integralLength > 0 ? sign + digits.Insert(integralLength, ".") :
            sign + "0." + new string('0', -integralLength) + digits;
    }

    private static Term ParseTerm(JsonElement element)
    {
        if (element.ValueKind != JsonValueKind.Number) throw new NotSupportedException("fill series operands must be numbers");
        var raw = element.GetRawText();
        if (raw.Length > 128) throw new NotSupportedException("fill series operand is too long");
        var match = NumberPattern.Match(raw);
        if (!match.Success || match.Groups["exponent"].Success &&
            !int.TryParse(match.Groups["exponent"].Value, NumberStyles.AllowLeadingSign,
                CultureInfo.InvariantCulture, out _))
            throw new NotSupportedException("fill series requires finite decimal operands");
        var exponent = match.Groups["exponent"].Success
            ? int.Parse(match.Groups["exponent"].Value, CultureInfo.InvariantCulture) : 0;
        var fraction = match.Groups["fraction"].Value;
        var digits = match.Groups["whole"].Value + fraction;
        if (!BigInteger.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var coefficient))
            throw new NotSupportedException("Invalid fill series operand");
        if (coefficient.IsZero) return new Term(BigInteger.Zero, 0);
        if (match.Groups["sign"].Value == "-") coefficient = -coefficient;
        var power = (long)exponent - fraction.Length;
        while (coefficient % 10 == 0)
        {
            coefficient /= 10;
            power++;
        }
        if (power is < -30 or > 14)
            throw new NotSupportedException("fill series operand exponent is outside the supported range");
        return new Term(coefficient, (int)power);
    }

    private readonly record struct Term(BigInteger Coefficient, int Power);
}
