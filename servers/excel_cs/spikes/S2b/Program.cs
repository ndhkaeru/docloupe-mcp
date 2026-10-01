using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using DocLoupe.Excel.Verify;

if (args is ["normalize-types", var normalizationRoot, var normalizationOutput])
{
    SyntheticFixtures.NormalizeContentTypes(normalizationRoot, normalizationOutput);
    return;
}

if (args.Length == 3 && args[0] == "p1")
{
    var readRoot = args[1];
    var invalidPath = args[2];
    var readFiles = Directory.GetFiles(readRoot).Where(path => Path.GetFileName(path).Length > 2
        && Path.GetFileName(path)[..2] is "00" or "01" or "02" or "03" or "04" or "05" or "06" or "07")
        .Where(path => !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal)).OrderBy(path => path).ToArray();
    var readCount = 0;
    foreach (var path in readFiles)
    {
        try
        {
            var preview = WorkbookReader.Peek(path);
            var verification = WorkbookReader.VerifyPartial(path);
            if (preview.Sheets.Count > 0 && verification.Status == "unverified") readCount++;
            Console.WriteLine($"{Path.GetFileName(path)} | sheets={preview.Sheets.Count} cells={preview.FirstSheetCells.Count} status={verification.Status} g1_issues={verification.PackageIssues.Count} g3_issues={verification.MarkupIssues.Count}");
        }
        catch (Exception exception)
        {
            Console.WriteLine($"{Path.GetFileName(path)} | ERROR {exception.GetType().Name}: {exception.Message}");
        }
    }
    var invalidVerification = WorkbookReader.VerifyPartial(invalidPath);
    Console.WriteLine($"EX04_invalid.xlsx | status={invalidVerification.Status} g3={string.Join(',', invalidVerification.MarkupIssues.Select(issue => issue.Code))}");
    if (invalidVerification.Status != "failed" || !invalidVerification.MarkupIssues.Any(issue => issue.Code == "UNDECLARED_MC_PREFIX"))
        Environment.ExitCode = 1;
    Console.WriteLine($"P1_PARTIAL | read={readCount}/{readFiles.Length}");
    if (readCount != readFiles.Length) Environment.ExitCode = 1;
    return;
}

if (args is not ["probe", var sourceRoot, var outputRoot])
    throw new ArgumentException("Usage: probe|normalize-types <immutable-fixtures-dir> <output-dir>");

Directory.CreateDirectory(outputRoot);
var fixtures = Directory.GetFiles(sourceRoot)
    .Where(path => Path.GetFileName(path).Length > 2 && Path.GetFileName(path)[..2] is "00" or "01" or "02" or "03" or "04" or "05" or "06" or "07")
    .Where(path => !Path.GetFileName(path).StartsWith("07-external-", StringComparison.Ordinal))
    .OrderBy(path => path).ToArray();
if (fixtures.Length != 8) throw new InvalidOperationException($"Expected 8 fixtures, found {fixtures.Length}");

var successes = 0;
foreach (var source in fixtures)
{
    var sourceHash = SHA256.HashData(File.ReadAllBytes(source));
    var output = Path.Combine(outputRoot, Path.GetFileNameWithoutExtension(source) + "-s2b" + Path.GetExtension(source));
    var result = RawWorksheetEditor.Edit(source, output);
    var sourceUnchanged = sourceHash.AsSpan().SequenceEqual(SHA256.HashData(File.ReadAllBytes(source)));
    var passed = sourceUnchanged && result.G3.Passed && result.DeclarationPreserved && result.RootPrefixPreserved
        && result.RootNamespacesPreserved && result.UntouchedPartsPreserved && result.UntouchedSheetBytesPreserved && result.EditedCellPresent;
    if (passed) successes++;
    Console.WriteLine($"{Path.GetFileName(source)} | typed_cells={result.TypedCellCount} parts={result.ParsedPartCount} cell={result.CellReference} | source={(sourceUnchanged ? "same" : "CHANGED")} untouched={(result.UntouchedPartsPreserved ? "same" : "CHANGED")} outside_cell={(result.UntouchedSheetBytesPreserved ? "same" : "CHANGED")} prolog={(result.DeclarationPreserved ? "same" : "CHANGED")} prefixes={(result.RootPrefixPreserved ? "same" : "CHANGED")} declarations={(result.RootNamespacesPreserved ? "same" : "CHANGED")} g3={(result.G3.Passed ? "pass" : string.Join(',', result.G3.Issues.Select(issue => issue.Code)))} edit={(result.EditedCellPresent ? "ok" : "LOST")}");
}

var synthetic = Path.Combine(outputRoot, "V06_EX03_valid.xlsx");
SyntheticFixtures.Create(fixtures[0], synthetic, false);
var syntheticOutput = Path.Combine(outputRoot, "V06_EX03_valid-s2b.xlsx");
var syntheticResult = RawWorksheetEditor.Edit(synthetic, syntheticOutput);
Console.WriteLine($"V06_EX03_valid.xlsx | outside_cell={syntheticResult.UntouchedSheetBytesPreserved} prolog={syntheticResult.DeclarationPreserved} prefixes={syntheticResult.RootPrefixPreserved} declarations={syntheticResult.RootNamespacesPreserved} g3={syntheticResult.G3.Passed} edit={syntheticResult.EditedCellPresent}");

var invalid = Path.Combine(outputRoot, "EX04_invalid.xlsx");
SyntheticFixtures.Create(fixtures[0], invalid, true);
using var invalidZip = ZipFile.OpenRead(invalid);
using var invalidXml = invalidZip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
var invalidResult = MarkupCompatibilityVerifier.Check(invalidXml);
Console.WriteLine($"EX04_invalid.xlsx | g3={(invalidResult.Passed ? "INCORRECT_PASS" : string.Join(',', invalidResult.Issues.Select(issue => issue.Code)))}");

var ex03 = Path.Combine(outputRoot, "EX03_namespace_loss.xlsx");
SyntheticFixtures.CorruptX14acNamespace(synthetic, ex03);
using var ex03Zip = ZipFile.OpenRead(ex03);
using var validZip = ZipFile.OpenRead(synthetic);
using var ex03Xml = ex03Zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
using var validXml = validZip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
var ex03Result = MarkupCompatibilityVerifier.Check(ex03Xml, validXml);
Console.WriteLine($"EX03_namespace_loss.xlsx | g3={(ex03Result.Passed ? "INCORRECT_PASS" : string.Join(',', ex03Result.Issues.Select(issue => issue.Code)))}");

var pass = successes == 8 && syntheticResult.G3.Passed && syntheticResult.DeclarationPreserved
    && syntheticResult.RootPrefixPreserved && syntheticResult.RootNamespacesPreserved && syntheticResult.UntouchedSheetBytesPreserved && syntheticResult.EditedCellPresent
    && !invalidResult.Passed && invalidResult.Issues.Any(issue => issue.Code == "UNDECLARED_MC_PREFIX")
    && ex03Result.Issues.Any(issue => issue.Code == "ATTRIBUTE_NAMESPACE_CHANGED");
Console.WriteLine($"SUMMARY | fixtures={successes}/8 | synthetic={(syntheticResult.G3.Passed ? "pass" : "fail")} | EX04_rejected={!invalidResult.Passed} EX03_rejected={!ex03Result.Passed} | result={(pass ? "PASS" : "FAIL")}");
if (!pass) Environment.ExitCode = 1;
