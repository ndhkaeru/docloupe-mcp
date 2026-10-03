using System.IO.Compression;
using System.Text;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Verify.Tests;

public sealed class PackageComparatorTests
{
    [Fact]
    public void IgnoresZipContainerChangesButComparesDecompressedPartBytes()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-compare-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var before = Path.Combine(directory, "before.xlsx");
            var after = Path.Combine(directory, "after.xlsx");
            Create(before, CompressionLevel.NoCompression, ("xl/sheet.xml", "same"));
            Create(after, CompressionLevel.Optimal, ("XL/SHEET.XML", "same"));
            var result = PackageComparator.Compare(before, after);
            Assert.False(result.HasDifferences);
            Assert.False(result.Truncated);
            Assert.Empty(result.Differences);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ReportsAddedRemovedAndChangedPartsWithStableIdsAndBoundedOutput()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-compare-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var before = Path.Combine(directory, "before.xlsx");
            var after = Path.Combine(directory, "after.xlsx");
            Create(before, CompressionLevel.Optimal, ("xl/changed.xml", "old"), ("xl/removed.xml", "gone"));
            Create(after, CompressionLevel.Optimal, ("xl/added.xml", "new"), ("xl/changed.xml", "changed"));
            var result = PackageComparator.Compare(before, after);
            Assert.True(result.HasDifferences);
            Assert.False(result.Truncated);
            Assert.Equal(new[] { "added_part", "changed_part", "removed_part" },
                result.Differences.Select(difference => difference.Category));
            Assert.All(result.Differences, difference => Assert.StartsWith("d_", difference.Id));
            Assert.Equal(result.Differences.Select(difference => difference.Id),
                PackageComparator.Compare(before, after).Differences.Select(difference => difference.Id));
            Assert.True(PackageComparator.Compare(before, after, 1).Truncated);
            Assert.Single(PackageComparator.Compare(before, after, 1).Differences);
            Assert.False(PackageComparator.Compare(before, after, 3).Truncated);
            Assert.Throws<ArgumentOutOfRangeException>(() => PackageComparator.Compare(before, after, 0));
        }
        finally { Directory.Delete(directory, true); }
    }

    private static void Create(string path, CompressionLevel compression, params (string Name, string Value)[] parts)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var (name, value) in parts)
        {
            using var stream = archive.CreateEntry(name, compression).Open();
            stream.Write(Encoding.UTF8.GetBytes(value));
        }
    }
}
