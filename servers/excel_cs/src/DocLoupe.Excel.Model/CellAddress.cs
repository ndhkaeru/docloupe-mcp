using System.Globalization;

namespace DocLoupe.Excel.Model;

public readonly record struct CellAddress(int Row, int Column)
{
    public static string? SheetName(string address)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);
        var separator = address.LastIndexOf('!');
        if (separator < 0) return null;
        var name = address[..separator];
        if (name.StartsWith('\'') && name.EndsWith('\'') && name.Length >= 3)
        {
            name = name[1..^1];
            for (var index = 0; index < name.Length; index++)
                if (name[index] == '\'' && (index + 1 >= name.Length || name[++index] != '\''))
                    throw new FormatException("Invalid quoted sheet name");
            return name.Replace("''", "'", StringComparison.Ordinal);
        }
        if (name.Length == 0 || name.Contains('\'')) throw new FormatException("Invalid sheet name");
        return name;
    }

    public static CellAddress Parse(string address)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);
        var separator = address.LastIndexOf('!');
        var token = separator < 0 ? address : address[(separator + 1)..];
        var position = 0;
        var column = 0;
        while (position < token.Length && token[position] is >= 'A' and <= 'Z' or >= 'a' and <= 'z')
        {
            column = checked(column * 26 + char.ToUpperInvariant(token[position]) - 'A' + 1);
            position++;
        }
        if (column is < 1 or > 16384 || position == 0 || position == token.Length
            || token[position] == '0' || !int.TryParse(token.AsSpan(position), NumberStyles.None, CultureInfo.InvariantCulture, out var row)
            || row is < 1 or > 1048576)
            throw new FormatException($"Invalid A1 cell address: {address}");
        return new CellAddress(row, column);
    }

    public override string ToString()
    {
        var column = Column;
        var letters = "";
        while (column > 0)
        {
            column--;
            letters = (char)('A' + column % 26) + letters;
            column /= 26;
        }
        return letters + Row.ToString(CultureInfo.InvariantCulture);
    }
}
