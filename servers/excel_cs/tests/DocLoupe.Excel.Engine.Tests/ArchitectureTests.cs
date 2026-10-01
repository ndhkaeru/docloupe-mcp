using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ArchitectureTests
{
    [Fact]
    public void VerifierDoesNotReferenceEngineSchemaOrOpenXmlSdk()
    {
        using var stream = File.OpenRead(typeof(P2aGates).Assembly.Location);
        using var file = new PEReader(stream);
        var metadata = file.GetMetadataReader();
        var names = metadata.AssemblyReferences.Select(reference => metadata.GetString(metadata.GetAssemblyReference(reference).Name)).ToArray();
        Assert.DoesNotContain(names, name => name is "DocLoupe.Excel.Engine" or "DocLoupe.Excel.Schema"
            || name.StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal));
    }

    [Fact]
    public void EngineAndPackageDoNotReferenceOpenXmlSdk()
    {
        foreach (var assembly in new[] { typeof(DocLoupe.Excel.Engine.SetValueEngine).Assembly, typeof(DocLoupe.Excel.Package.PackageStore).Assembly })
        {
            using var stream = File.OpenRead(assembly.Location);
            using var file = new PEReader(stream);
            var metadata = file.GetMetadataReader();
            Assert.DoesNotContain(metadata.AssemblyReferences, reference =>
                metadata.GetString(metadata.GetAssemblyReference(reference).Name).StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal));
        }
    }
}
