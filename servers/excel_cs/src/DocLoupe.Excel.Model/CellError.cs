namespace DocLoupe.Excel.Model;

public static class CellError
{
    private static readonly HashSet<string> Tokens = new(StringComparer.Ordinal)
    {
        "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!",
        "#GETTING_DATA", "#SPILL!", "#CALC!", "#BLOCKED!", "#CONNECT!",
        "#EXTERNAL!", "#FIELD!", "#UNKNOWN!"
    };

    public static bool IsSupported(string? value) => value is not null && Tokens.Contains(value);
}
