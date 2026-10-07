using System.Text;
using System.Xml;
using DocLoupe.Excel.Package;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

internal static class StyledFixture
{
    public static void Create(string source, string destination)
    {
        using var store = new PackageStore(source);
        var relationshipsPart = PackageStore.RelationshipPart(store.WorkbookPart);
        var relationships = PackageStore.Parse(store.Read(relationshipsPart));
        var relationship = relationships.CreateElement("Relationship", PackageStore.PackageRelationships);
        relationship.SetAttribute("Id", "rId3");
        relationship.SetAttribute("Type", PackageStore.OfficeRelationships + "/styles");
        relationship.SetAttribute("Target", "styles.xml");
        relationships.DocumentElement!.AppendChild(relationship);
        store.Set(relationshipsPart, Encoding.UTF8.GetBytes(relationships.OuterXml));

        var contentTypes = PackageStore.Parse(store.Read("[Content_Types].xml"));
        var contentType = contentTypes.CreateElement("Override", PackageStore.ContentTypes);
        contentType.SetAttribute("PartName", "/xl/styles.xml");
        contentType.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        contentTypes.DocumentElement!.AppendChild(contentType);
        store.Set("[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes.OuterXml));

        const string styles = """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>
            """;
        store.Set("xl/styles.xml", Encoding.UTF8.GetBytes(styles));
        var sheetPart = store.SheetPart("Sheet1");
        var sheet = PackageStore.Parse(store.Read(sheetPart));
        foreach (var (address, styleIndex) in new[] { ("B1", "1"), ("D3", "0") })
        {
            var cell = Assert.Single(sheet.GetElementsByTagName("c", PackageStore.Main).OfType<XmlElement>(),
                element => element.GetAttribute("r") == address);
            cell.SetAttribute("s", styleIndex);
        }
        store.Set(sheetPart, Encoding.UTF8.GetBytes(sheet.OuterXml));
        store.Save(destination);
    }
}
