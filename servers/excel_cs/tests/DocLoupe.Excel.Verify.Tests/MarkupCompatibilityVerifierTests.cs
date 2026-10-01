using System.Text;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Verify.Tests;

public sealed class MarkupCompatibilityVerifierTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Markup = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string X14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private static readonly string Valid = $"<x:worksheet xmlns:x='{Main}' xmlns:mc='{Markup}' xmlns:x14ac='{X14ac}' mc:Ignorable='x14ac'><x:sheetFormatPr x14ac:dyDescent='0.25'/><x:sheetData><x:row r='1'><x:c r='A1' t='inlineStr'><x:is><x:t>before</x:t></x:is></x:c><x:c r='B1'><x:v>10</x:v></x:c></x:row></x:sheetData></x:worksheet>";

    [Fact]
    public void DetectsUndeclaredIgnorablePrefix()
    {
        var xml = Valid.Replace("mc:Ignorable='x14ac'", "mc:Ignorable='x14ac missing'", StringComparison.Ordinal);
        Assert.Contains(Check(xml).Issues, issue => issue.Code == "UNDECLARED_MC_PREFIX");
    }

    [Fact]
    public void DetectsLossOfAttributeNamespace()
    {
        var xml = Valid.Replace("x14ac:dyDescent", "dyDescent", StringComparison.Ordinal);
        Assert.Contains(Check(xml, Valid).Issues, issue => issue.Code == "ATTRIBUTE_NAMESPACE_CHANGED");
    }

    [Fact]
    public void AcceptsAnEditWithoutSkippingLaterCells()
    {
        var xml = Valid.Replace("before", "after", StringComparison.Ordinal);
        Assert.True(Check(xml, Valid, "A1").Passed);
        var changed = xml.Replace("<x:c r='B1'>", "<x:c r='B1' xmlns:z='urn:other' z:tag='1'>", StringComparison.Ordinal);
        Assert.Contains(Check(changed, Valid, "A1").Issues, issue => issue.Code == "ATTRIBUTE_NAMESPACE_CHANGED");
    }

    [Fact]
    public void RejectsChangedPrefixesAndRootDeclarations()
    {
        var xml = Valid.Replace("xmlns:x=", "xmlns:y=", StringComparison.Ordinal).Replace("x:", "y:", StringComparison.Ordinal);
        var result = Check(xml, Valid, "A1");
        Assert.Contains(result.Issues, issue => issue.Code == "ROOT_NAMESPACE_CHANGED");
        Assert.Contains(result.Issues, issue => issue.Code == "ELEMENT_NAMESPACE_CHANGED");
    }

    [Fact]
    public void DetectsUndeclaredPrefixInChoiceRequires()
    {
        // mc:Choice/@Requires is unqualified, so it is easy to miss when only mc:-qualified attributes are checked.
        var xml = $"<x:worksheet xmlns:x='{Main}' xmlns:mc='{Markup}'><mc:AlternateContent><mc:Choice Requires='missing'><x:sheetData/></mc:Choice><mc:Fallback><x:sheetData/></mc:Fallback></mc:AlternateContent></x:worksheet>";
        Assert.Contains(Check(xml).Issues, issue => issue.Code == "UNDECLARED_MC_PREFIX");
        var declared = xml.Replace("xmlns:mc=", $"xmlns:missing='{X14ac}' xmlns:mc=", StringComparison.Ordinal);
        Assert.True(Check(declared).Passed);
    }

    [Fact]
    public void RejectsDtds()
    {
        var xml = "<!DOCTYPE worksheet [<!ENTITY x 'unsafe'>]>" + Valid;
        Assert.Contains(Check(xml).Issues, issue => issue.Code == "INVALID_XML");
    }

    private static MarkupResult Check(string after, string? before = null, string? editedCell = null)
    {
        using var afterStream = new MemoryStream(Encoding.UTF8.GetBytes(after));
        using var beforeStream = before is null ? null : new MemoryStream(Encoding.UTF8.GetBytes(before));
        return MarkupCompatibilityVerifier.Check(afterStream, beforeStream, editedCell);
    }
}
