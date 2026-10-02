using System.Text.Json;
using System.Text.Json.Serialization;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Model;

namespace DocLoupe.Excel.Server;

public sealed class SetValueRequest
{
    [JsonPropertyName("op")]
    public required string Op { get; init; }
    [JsonPropertyName("sheet")]
    public string? Sheet { get; init; }
    [JsonPropertyName("target")]
    public required string Target { get; init; }
    [JsonPropertyName("value")]
    public JsonElement Value { get; init; }
    [JsonPropertyName("values")]
    public JsonElement Values { get; init; }
    [JsonPropertyName("as_text")]
    public bool AsText { get; init; }
    [JsonPropertyName("rich_policy")]
    public string RichPolicy { get; init; } = "reject";
    [JsonPropertyName("formula")]
    public string? Formula { get; init; }
    [JsonPropertyName("kind")]
    public string? FormulaKind { get; init; }
    [JsonPropertyName("ref")]
    public string? Reference { get; init; }
    [JsonPropertyName("cache")]
    public JsonElement? Cache { get; init; }
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Other { get; init; }

    public SetValueOp Normalize(string? defaultSheet)
    {
        if (Op is not ("set_value" or "set_formula")) throw new NotSupportedException("Only set_value and normal set_formula are supported");
        if (Other is { Count: > 0 }) throw new NotSupportedException("Unsupported cell operation fields");
        if (Values.ValueKind != JsonValueKind.Undefined) throw new NotSupportedException("values requires set_values");
        var name = CellAddress.SheetName(Target) ?? Sheet ?? defaultSheet ?? throw new ArgumentException("Missing sheet name");
        var address = CellAddress.Parse(Target).ToString();
        if (Op == "set_formula")
        {
            if (Value.ValueKind != JsonValueKind.Undefined || AsText || RichPolicy != "reject" ||
                FormulaKind is not (null or "normal") || Reference is not null ||
                Cache is { } cache && (cache.ValueKind != JsonValueKind.String || cache.GetString() != "clear"))
                throw new NotSupportedException("Only normal set_formula with cleared cache is supported");
            var formula = Formula?.StartsWith('=') == true ? Formula[1..] : Formula;
            if (string.IsNullOrWhiteSpace(formula) || formula.StartsWith('='))
                throw new ArgumentException("Formula must be a nonempty expression");
            return new SetValueOp(name, address, "formula", formula, Operation: Op);
        }
        if (Formula is not null || FormulaKind is not null || Reference is not null || Cache is not null)
            throw new NotSupportedException("Formula fields require set_formula");
        var (kind, scalar) = Value.ValueKind switch
        {
            JsonValueKind.String => ("text", Value.GetString()),
            JsonValueKind.Number => ("number", Value.GetRawText()),
            JsonValueKind.True => ("boolean", "true"),
            JsonValueKind.False => ("boolean", "false"),
            JsonValueKind.Null => ("blank", (string?)null),
            JsonValueKind.Object when Value.TryGetProperty("formula", out var formula) && formula.ValueKind == JsonValueKind.String
                => ("formula", formula.GetString()),
            JsonValueKind.Object when Value.TryGetProperty("inline", out var inline) && inline.ValueKind == JsonValueKind.String
                => ("inline", inline.GetString()),
            _ => throw new ArgumentException("Unsupported set_value value")
        };
        if (AsText && kind != "text") throw new ArgumentException("as_text requires a string value");
        if (kind == "text" && scalar?.StartsWith('=') == true && !AsText)
            throw new ArgumentException("AMBIGUOUS_FORMULA_TEXT");
        return new SetValueOp(name, address, kind, scalar, RichPolicy, AsText, Op);
    }

    public SetValueOp[] NormalizeMany(string? defaultSheet)
    {
        if (Op == "fill" && (!Target.Contains(':') || Value.ValueKind == JsonValueKind.Undefined ||
            Values.ValueKind != JsonValueKind.Undefined || AsText || RichPolicy != "reject" ||
            Formula is not null || FormulaKind is not null || Reference is not null || Cache is not null ||
            Other is { Count: > 0 }))
            throw new NotSupportedException("Only constant fill on a rectangular range is supported");
        if ((Op is "set_value" or "fill") && Target.Contains(':'))
        {
            var (sheet, rangeStart, rangeEnd, _) = ParseTargetRange(defaultSheet);
            var count = (long)(rangeEnd.Row - rangeStart.Row + 1) * (rangeEnd.Column - rangeStart.Column + 1);
            if (count > 500) throw new ArgumentException($"{Op} range exceeds 500 cells");
            var broadcast = new List<SetValueOp>((int)count);
            for (var row = rangeStart.Row; row <= rangeEnd.Row; row++)
                for (var column = rangeStart.Column; column <= rangeEnd.Column; column++)
                    broadcast.Add(new SetValueRequest
                    {
                        Op = "set_value", Sheet = sheet, Target = new CellAddress(row, column).ToString(),
                        Value = Value, Values = Values, AsText = AsText, RichPolicy = RichPolicy,
                        Formula = Formula, FormulaKind = FormulaKind, Reference = Reference, Cache = Cache, Other = Other
                    }.Normalize(sheet) with { Operation = Op });
            return broadcast.ToArray();
        }
        if (Op != "set_values") return [Normalize(defaultSheet)];
        if (Value.ValueKind != JsonValueKind.Undefined || Values.ValueKind != JsonValueKind.Array ||
            AsText || RichPolicy != "reject" || Formula is not null || FormulaKind is not null ||
            Reference is not null || Cache is not null || Other is { Count: > 0 })
            throw new NotSupportedException("set_values requires only a target and rectangular values array");

        var (name, first, last, hasRange) = ParseTargetRange(defaultSheet);

        var rows = Values.EnumerateArray().ToArray();
        if (rows.Length is < 1 or > 500 || rows.Any(row => row.ValueKind != JsonValueKind.Array))
            throw new ArgumentException("set_values requires a nonempty two-dimensional array of at most 500 cells");
        var width = rows[0].GetArrayLength();
        if (width is < 1 or > 500 || rows.Length * width > 500 || rows.Any(row => row.GetArrayLength() != width))
            throw new ArgumentException("set_values requires a rectangular array of at most 500 cells");
        if (hasRange && (last.Row - first.Row + 1 != rows.Length || last.Column - first.Column + 1 != width))
            throw new ArgumentException("set_values dimensions do not match target range");
        if (first.Row + rows.Length - 1 > 1048576 || first.Column + width - 1 > 16384)
            throw new ArgumentException("set_values exceeds worksheet bounds");

        var operations = new List<SetValueOp>(rows.Length * width);
        for (var row = 0; row < rows.Length; row++)
            for (var column = 0; column < width; column++)
                operations.Add(new SetValueRequest
                {
                    Op = "set_value", Sheet = name,
                    Target = new CellAddress(first.Row + row, first.Column + column).ToString(),
                    Value = rows[row][column]
                }.Normalize(name) with { Operation = "set_values" });
        return operations.ToArray();
    }

    private (string Sheet, CellAddress First, CellAddress Last, bool HasRange) ParseTargetRange(string? defaultSheet)
    {
        var name = CellAddress.SheetName(Target) ?? Sheet ?? defaultSheet ?? throw new ArgumentException("Missing sheet name");
        var separator = Target.LastIndexOf('!');
        var bounds = (separator < 0 ? Target : Target[(separator + 1)..]).Split(':');
        if (bounds.Length is < 1 or > 2) throw new FormatException("Invalid cell range");
        var first = CellAddress.Parse(bounds[0]);
        var last = bounds.Length == 2 ? CellAddress.Parse(bounds[1]) : first;
        if (last.Row < first.Row || last.Column < first.Column) throw new FormatException("Reversed cell range");
        return (name, first, last, bounds.Length == 2);
    }
}

public sealed class SaveAssertionRequest
{
    [JsonPropertyName("target")]
    public required string Target { get; init; }
    [JsonPropertyName("equals")]
    public required JsonElement Expected { get; init; }
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Other { get; init; }

    public DocLoupe.Excel.Verify.ValueAssertion Normalize()
    {
        if (Other is { Count: > 0 }) throw new NotSupportedException("Only equals assertions are supported");
        var sheet = CellAddress.SheetName(Target) ?? throw new ArgumentException("Assertion target must be sheet-qualified");
        var address = CellAddress.Parse(Target).ToString();
        if (Expected.ValueKind != JsonValueKind.Object) throw new ArgumentException("Assertion equals must be an object");
        var properties = Expected.EnumerateObject().ToArray();
        if (properties.Length == 0 || properties.GroupBy(property => property.Name).Any(group => group.Count() > 1) ||
            properties.Any(property => property.Name is not ("value" or "formula")))
            throw new NotSupportedException("Only equals.value and equals.formula assertions are supported");
        var checkValue = Expected.TryGetProperty("value", out var value);
        var (kind, scalar) = checkValue ? value.ValueKind switch
        {
            JsonValueKind.String => ("text", value.GetString()),
            JsonValueKind.Number => ("number", value.GetRawText()),
            JsonValueKind.True => ("boolean", "true"),
            JsonValueKind.False => ("boolean", "false"),
            JsonValueKind.Null => ("blank", (string?)null),
            _ => throw new NotSupportedException("Unsupported assertion value")
        } : ((string?)null, (string?)null);
        if (kind == "number" && !DocLoupe.Excel.Verify.G7Assertions.IsSupportedNumber(scalar))
            throw new NotSupportedException("Numeric assertion is outside the supported lexical range");
        string? formula = null;
        if (Expected.TryGetProperty("formula", out var expectedFormula))
        {
            if (expectedFormula.ValueKind != JsonValueKind.String ||
                string.IsNullOrEmpty(formula = expectedFormula.GetString()?.TrimStart('=')))
                throw new ArgumentException("Assertion formula must be a nonempty string");
        }
        return new DocLoupe.Excel.Verify.ValueAssertion(sheet, address, checkValue, kind, scalar, formula);
    }
}
