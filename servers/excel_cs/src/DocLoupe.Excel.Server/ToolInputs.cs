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
    [JsonPropertyName("series")]
    public JsonElement Series { get; init; }
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
    [JsonPropertyName("what")]
    public JsonElement What { get; init; }
    [JsonPropertyName("remove_cells")]
    public JsonElement RemoveCells { get; init; }
    [JsonPropertyName("expect")]
    public JsonElement Expect { get; init; }
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Other { get; init; }

    public SetValueOp Normalize(string? defaultSheet)
    {
        if (Op is not ("set_value" or "set_formula" or "clear")) throw new NotSupportedException("Unsupported cell operation");
        if (Other is { Count: > 0 }) throw new NotSupportedException("Unsupported cell operation fields");
        if (Op == "clear")
        {
            if (Value.ValueKind != JsonValueKind.Undefined || Values.ValueKind != JsonValueKind.Undefined ||
                Series.ValueKind != JsonValueKind.Undefined || AsText || RichPolicy != "reject" ||
                Formula is not null || FormulaKind is not null || Reference is not null || Cache is not null ||
                RemoveCells.ValueKind is not (JsonValueKind.Undefined or JsonValueKind.False or JsonValueKind.True) ||
                What.ValueKind != JsonValueKind.Undefined &&
                (What.ValueKind != JsonValueKind.Array || What.GetArrayLength() != 1 ||
                 What[0].ValueKind != JsonValueKind.String || What[0].GetString() != "values"))
                throw new NotSupportedException("Only clear values with optional remove_cells is supported");
            var clearSheet = CellAddress.SheetName(Target) ?? Sheet ?? defaultSheet ?? throw new ArgumentException("Missing sheet name");
            var clearAddress = CellAddress.Parse(Target).ToString();
            return new SetValueOp(clearSheet, clearAddress, "blank", null,
                Operation: "clear", RemoveCell: RemoveCells.ValueKind == JsonValueKind.True,
                Expect: NormalizeExpect(clearSheet, clearAddress));
        }
        if (What.ValueKind != JsonValueKind.Undefined || RemoveCells.ValueKind != JsonValueKind.Undefined)
            throw new NotSupportedException("what and remove_cells require clear");
        if (Values.ValueKind != JsonValueKind.Undefined) throw new NotSupportedException("values requires set_values");
        if (Series.ValueKind != JsonValueKind.Undefined) throw new NotSupportedException("series requires fill");
        var name = CellAddress.SheetName(Target) ?? Sheet ?? defaultSheet ?? throw new ArgumentException("Missing sheet name");
        var address = CellAddress.Parse(Target).ToString();
        if (Op == "set_formula")
        {
            if (Value.ValueKind != JsonValueKind.Undefined || AsText || RichPolicy != "reject" ||
                FormulaKind is not (null or "normal") || Reference is not null ||
                Cache is { } cache && cache.ValueKind != JsonValueKind.Object &&
                (cache.ValueKind != JsonValueKind.String || cache.GetString() is not ("clear" or "keep")))
                throw new NotSupportedException("Only normal set_formula with clear, keep, or an explicit cache is supported");
            var formula = Formula?.StartsWith('=') == true ? Formula[1..] : Formula;
            if (string.IsNullOrWhiteSpace(formula) || formula.StartsWith('='))
                throw new ArgumentException("Formula must be a nonempty expression");
            FormulaCache? explicitCache = null;
            if (Cache is { ValueKind: JsonValueKind.Object } cacheObject)
            {
                if (cacheObject.EnumerateObject().Count() != 1 ||
                    !cacheObject.TryGetProperty("value", out var cached))
                    throw new ArgumentException("cache requires exactly one value field");
                explicitCache = cached.ValueKind switch
                {
                    JsonValueKind.Number => new FormulaCache("n", cached.GetRawText()),
                    JsonValueKind.String => new FormulaCache("str", cached.GetString()!),
                    JsonValueKind.True => new FormulaCache("b", "1"),
                    JsonValueKind.False => new FormulaCache("b", "0"),
                    JsonValueKind.Object when cached.EnumerateObject().Count() == 1 &&
                        cached.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.String &&
                        CellError.IsSupported(error.GetString()) => new FormulaCache("e", error.GetString()!),
                    _ => throw new ArgumentException("Unsupported formula cache value")
                };
            }
            return new SetValueOp(name, address, "formula", formula, Operation: Op,
                KeepCache: Cache is { ValueKind: JsonValueKind.String } policy && policy.GetString() == "keep",
                ExplicitCache: explicitCache, Expect: NormalizeExpect(name, address));
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
            JsonValueKind.Object when Value.EnumerateObject().Count() == 1 &&
                Value.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.String &&
                CellError.IsSupported(error.GetString()) => ("error", error.GetString()),
            _ => throw new ArgumentException("Unsupported set_value value")
        };
        if (AsText && kind != "text") throw new ArgumentException("as_text requires a string value");
        if (kind == "text" && scalar?.StartsWith('=') == true && !AsText)
            throw new ArgumentException("AMBIGUOUS_FORMULA_TEXT");
        return new SetValueOp(name, address, kind, scalar, RichPolicy, AsText, Op,
            Expect: NormalizeExpect(name, address));
    }

    private CellPrecondition? NormalizeExpect(string sheet, string address)
    {
        if (Expect.ValueKind == JsonValueKind.Undefined) return null;
        if (Expect.ValueKind != JsonValueKind.Object)
            throw new NotSupportedException("Only expect.value and expect.formula are supported");
        var assertion = new SaveAssertionRequest { Target = sheet + "!" + address, Expected = Expect }.Normalize();
        return new CellPrecondition(assertion.CheckValue, assertion.Kind, assertion.Value, assertion.Formula);
    }

    public SetValueOp[] NormalizeMany(string? defaultSheet)
    {
        if (Expect.ValueKind != JsonValueKind.Undefined && (Op is not ("set_value" or "set_formula" or "clear") || Target.Contains(':')))
            throw new NotSupportedException("expect requires a single-cell operation");
        var hasValue = Value.ValueKind != JsonValueKind.Undefined;
        var hasSeries = Series.ValueKind != JsonValueKind.Undefined;
        if (Op == "set_value" && hasSeries) throw new NotSupportedException("series requires fill");
        if (Op == "fill" && (!Target.Contains(':') || hasValue == hasSeries ||
            What.ValueKind != JsonValueKind.Undefined || RemoveCells.ValueKind != JsonValueKind.Undefined ||
            Values.ValueKind != JsonValueKind.Undefined || AsText || RichPolicy != "reject" ||
            Formula is not null || FormulaKind is not null || Reference is not null || Cache is not null ||
            Other is { Count: > 0 }))
            throw new NotSupportedException("fill requires exactly one of value or a numeric series on a rectangular range");
        if ((Op is "set_value" or "fill" or "clear") && Target.Contains(':'))
        {
            var (sheet, rangeStart, rangeEnd, _) = ParseTargetRange(defaultSheet);
            var count = (long)(rangeEnd.Row - rangeStart.Row + 1) * (rangeEnd.Column - rangeStart.Column + 1);
            if (count > 500) throw new ArgumentException($"{Op} range exceeds 500 cells");
            var series = hasSeries ? ExactSeries.Parse(Series) : null;
            var broadcast = new List<SetValueOp>((int)count);
            for (var row = rangeStart.Row; row <= rangeEnd.Row; row++)
                for (var column = rangeStart.Column; column <= rangeEnd.Column; column++)
                {
                    var address = new CellAddress(row, column).ToString();
                    if (Op == "clear")
                        broadcast.Add(new SetValueRequest
                        {
                            Op = "clear", Sheet = sheet, Target = address, What = What,
                            RemoveCells = RemoveCells, Other = Other
                        }.Normalize(sheet));
                    else if (series is not null)
                    {
                        var index = (row - rangeStart.Row) * (rangeEnd.Column - rangeStart.Column + 1) + column - rangeStart.Column;
                        broadcast.Add(new SetValueOp(sheet, address, "number", series.At(index), Operation: "fill"));
                    }
                    else
                        broadcast.Add(new SetValueRequest
                        {
                            Op = "set_value", Sheet = sheet, Target = address,
                            Value = Value, Values = Values, Series = Series, AsText = AsText, RichPolicy = RichPolicy,
                            Formula = Formula, FormulaKind = FormulaKind, Reference = Reference, Cache = Cache,
                            What = What, RemoveCells = RemoveCells, Other = Other
                        }.Normalize(sheet) with { Operation = Op });
                }
            return broadcast.ToArray();
        }
        if (Op != "set_values") return [Normalize(defaultSheet)];
        if (Value.ValueKind != JsonValueKind.Undefined || Values.ValueKind != JsonValueKind.Array ||
            Series.ValueKind != JsonValueKind.Undefined ||
            AsText || RichPolicy != "reject" || Formula is not null || FormulaKind is not null ||
            Reference is not null || Cache is not null || What.ValueKind != JsonValueKind.Undefined ||
            RemoveCells.ValueKind != JsonValueKind.Undefined || Other is { Count: > 0 })
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

public sealed class FindQueryRequest
{
    [JsonPropertyName("text")]
    public string? Text { get; init; }
    [JsonPropertyName("regex")]
    public string? Regex { get; init; }
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Other { get; init; }

    public (string Pattern, bool IsRegex) Normalize()
    {
        if (Other is { Count: > 0 } || (Text is null) == (Regex is null))
            throw new NotSupportedException("Only one of query.text or query.regex is supported");
        var pattern = Text ?? Regex!;
        if (pattern.Length is < 1 or > 512) throw new ArgumentException("Search pattern must contain 1..512 characters");
        return (pattern, Regex is not null);
    }
}

public sealed class FindScopeRequest
{
    [JsonPropertyName("sheet")]
    public string? Sheet { get; init; }
    [JsonPropertyName("target")]
    public required string Target { get; init; }
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Other { get; init; }
}

public sealed class SaveAssertionRequest
{
    [JsonPropertyName("target")]
    public required string Target { get; init; }
    [JsonPropertyName("equals")]
    public JsonElement Expected { get; init; }
    [JsonPropertyName("unchanged")]
    public JsonElement Unchanged { get; init; }
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Other { get; init; }

    public DocLoupe.Excel.Verify.ValueAssertion Normalize()
    {
        if (Other is { Count: > 0 }) throw new NotSupportedException("Unsupported assertion fields");
        var sheet = CellAddress.SheetName(Target) ?? throw new ArgumentException("Assertion target must be sheet-qualified");
        var address = CellAddress.Parse(Target).ToString();
        if (Unchanged.ValueKind != JsonValueKind.Undefined)
        {
            if (Unchanged.ValueKind != JsonValueKind.True || Expected.ValueKind != JsonValueKind.Undefined)
                throw new NotSupportedException("Only standalone unchanged: true is supported");
            return new DocLoupe.Excel.Verify.ValueAssertion(sheet, address, false, null, null, null, Unchanged: true);
        }
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
            JsonValueKind.Object when value.EnumerateObject().Count() == 1 &&
                value.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.String &&
                CellError.IsSupported(error.GetString()) => ("error", error.GetString()),
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
