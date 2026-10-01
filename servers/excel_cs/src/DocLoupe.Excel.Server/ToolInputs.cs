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
    [JsonPropertyName("as_text")]
    public bool AsText { get; init; }
    [JsonPropertyName("rich_policy")]
    public string RichPolicy { get; init; } = "reject";

    public SetValueOp Normalize(string? defaultSheet)
    {
        if (Op != "set_value") throw new NotSupportedException("P2a supports only set_value");
        var name = CellAddress.SheetName(Target) ?? Sheet ?? defaultSheet ?? throw new ArgumentException("Missing sheet name");
        var address = CellAddress.Parse(Target).ToString();
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
        return new SetValueOp(name, address, kind, scalar, RichPolicy, AsText);
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
