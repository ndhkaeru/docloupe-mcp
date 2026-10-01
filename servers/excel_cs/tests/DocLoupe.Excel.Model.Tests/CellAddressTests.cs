using DocLoupe.Excel.Model;
using Xunit;

namespace DocLoupe.Excel.Model.Tests;

public sealed class CellAddressTests
{
    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("Sheet1!XFD1048576", 1048576, 16384)]
    [InlineData("aa42", 42, 27)]
    public void ParsesA1(string input, int row, int column)
    {
        var address = CellAddress.Parse(input);
        Assert.Equal((row, column), (address.Row, address.Column));
        Assert.Equal(input.Split('!')[^1].ToUpperInvariant(), address.ToString());
    }

    [Theory]
    [InlineData("A0")]
    [InlineData("A01")]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("A1:B2")]
    public void RejectsInvalidCells(string input) => Assert.ThrowsAny<Exception>(() => CellAddress.Parse(input));
}
