using System.Text;
using System.Xml;
using DocLoupe.Excel.Package;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

internal static class RichFixture
{
    public static void Create(string source, string destination, bool inline)
    {
        using var store = new PackageStore(source);
        var part = inline ? store.SheetPart("Sheet1") : "xl/sharedStrings.xml";
        var document = PackageStore.Parse(store.Read(part));
        var container = inline
            ? Assert.Single(Assert.Single(document.GetElementsByTagName("c", PackageStore.Main)
                .OfType<XmlElement>(), cell => cell.GetAttribute("r") == "A1")
                .GetElementsByTagName("is", PackageStore.Main).OfType<XmlElement>())
            : Assert.Single(document.GetElementsByTagName("si", PackageStore.Main).OfType<XmlElement>());
        var originalText = Assert.Single(container.ChildNodes.OfType<XmlElement>(),
            node => node.LocalName == "t" && node.NamespaceURI == PackageStore.Main);
        container.RemoveChild(originalText);
        var plain = document.CreateElement(container.Prefix, "r", PackageStore.Main);
        plain.AppendChild(originalText);
        container.PrependChild(plain);
        var styled = document.CreateElement(container.Prefix, "r", PackageStore.Main);
        var properties = document.CreateElement(container.Prefix, "rPr", PackageStore.Main);
        var bold = document.CreateElement(container.Prefix, "b", PackageStore.Main);
        bold.SetAttribute("val", "1");
        properties.AppendChild(bold);
        var color = document.CreateElement(container.Prefix, "color", PackageStore.Main);
        color.SetAttribute("rgb", "FFFF0000");
        properties.AppendChild(color);
        styled.AppendChild(properties);
        var styledText = document.CreateElement(container.Prefix, "t", PackageStore.Main);
        var space = document.CreateAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace");
        space.Value = "preserve";
        styledText.Attributes.Append(space);
        styledText.InnerText = " bold";
        styled.AppendChild(styledText);
        container.InsertAfter(styled, plain);
        store.Set(part, Encoding.UTF8.GetBytes(document.OuterXml));
        store.Save(destination);
    }
}
