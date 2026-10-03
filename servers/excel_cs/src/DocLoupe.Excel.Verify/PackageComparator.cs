using System.IO.Compression;
using System.Security.Cryptography;

namespace DocLoupe.Excel.Verify;

public sealed record PartDifference(string Id, string Part, string Category, string? BeforeSha256,
    string? AfterSha256, string Classification = "undeclared");
public sealed record PackageComparison(IReadOnlyList<PartDifference> Differences, bool Truncated, bool HasDifferences);

public static class PackageComparator
{
    public static PackageComparison Compare(string beforePath, string afterPath, int maxDifferences = 200)
    {
        if (maxDifferences is < 1 or > 5000) throw new ArgumentOutOfRangeException(nameof(maxDifferences));
        using var before = ZipFile.OpenRead(beforePath);
        using var after = ZipFile.OpenRead(afterPath);
        var oldParts = before.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var newParts = after.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var names = oldParts.Keys.Union(newParts.Keys, StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
        var differences = new List<PartDifference>();
        var totalDifferences = 0;
        foreach (var name in names)
        {
            oldParts.TryGetValue(name, out var oldEntry);
            newParts.TryGetValue(name, out var newEntry);
            var oldHash = oldEntry is null ? null : Hash(oldEntry);
            var newHash = newEntry is null ? null : Hash(newEntry);
            if (oldHash == newHash) continue;
            totalDifferences++;
            if (differences.Count >= maxDifferences) continue;
            var category = oldEntry is null ? "added_part" : newEntry is null ? "removed_part" : "changed_part";
            var id = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(
                name + "\n" + (oldHash ?? "") + "\n" + (newHash ?? "")))).ToLowerInvariant();
            differences.Add(new PartDifference("d_" + id, name, category, oldHash, newHash));
        }
        return new PackageComparison(differences, totalDifferences > maxDifferences, totalDifferences > 0);
    }

    private static string Hash(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }
}
