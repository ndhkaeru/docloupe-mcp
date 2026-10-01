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
        var name = Sheet ?? CellAddress.SheetName(Target) ?? defaultSheet ?? throw new ArgumentException("Missing sheet name");
        CellAddress.Parse(Target);
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
        if (kind == "text" && scalar?.StartsWith('=') == true && !AsText)
            throw new ArgumentException("AMBIGUOUS_FORMULA_TEXT");
        if (kind == "text" && AsText && scalar?.StartsWith('=') == true) kind = "inline";
        return new SetValueOp(name, Target, kind, scalar, RichPolicy);
    }
}
