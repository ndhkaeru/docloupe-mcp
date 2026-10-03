using System.IO.Compression;
using DocLoupe.Excel.Package;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class PackageStoreTests
{
    [Theory]
    [InlineData("missing-root")]
    [InlineData("duplicate-part")]
    [InlineData("invalid-relationships")]
    public void FailedConstructionClosesSourceArchive(string scenario)
    {
        var path = Path.Combine(Path.GetTempPath(), "docloupe-package-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var archive = new ZipArchive(File.Create(path), ZipArchiveMode.Create))
            {
                if (scenario != "missing-root")
                {
                    Write(archive, "[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>");
                    Write(archive, "_rels/.rels", "<broken/>");
                }
                if (scenario == "duplicate-part")
                    Write(archive, "[content_types].XML", "<broken/>");
            }
            Assert.Throws<InvalidDataException>(() => new PackageStore(path));
            using var exclusive = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        finally { File.Delete(path); }
    }

    private static void Write(ZipArchive archive, string part, string value)
    {
        using var writer = new StreamWriter(archive.CreateEntry(part).Open());
        writer.Write(value);
    }
}
