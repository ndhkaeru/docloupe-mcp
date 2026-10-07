using System.Text;
using System.Text.Json;
using DocLoupe.Excel.Engine;
using DocLoupe.Excel.Package;
using DocLoupe.Excel.Server;
using DocLoupe.Excel.Verify;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class SaveAssertionTests
{
    [Fact]
    public void PerCellPreconditionMatcherAgreesWithSaveGate()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-consistency-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            var assertions = new[]
            {
                new ValueAssertion("Sheet1", "A1", true, "text", "hello", null),
                new ValueAssertion("Sheet1", "B1", true, "number", "42.0", null),
                new ValueAssertion("Sheet1", "B1", true, "number", "13", null),
                new ValueAssertion("Sheet1", "C1", true, "number", "2", "1+1"),
                new ValueAssertion("Sheet1", "C1", true, "number", "13", "9+9"),
                new ValueAssertion("Sheet1", "D3", true, "text", "old", null),
                new ValueAssertion("Sheet1", "E5", true, "blank", null, null),
                new ValueAssertion("Sheet1", "E5", false, null, null, "9+9")
            };
            var cells = P2aGates.ReadCells(source, "Sheet1", assertions.Select(item => item.Address))
                .ToDictionary(cell => cell.Address, StringComparer.Ordinal);
            foreach (var assertion in assertions)
                Assert.Equal(G7Assertions.Check(source, [assertion]).Count == 0,
                    G7Assertions.Matches(cells.GetValueOrDefault(assertion.Address), assertion));
            Assert.Equal(2, G7Assertions.Check(source, [assertions[4]]).Count);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void G7MatchesValuesAndFormulasIndependentlyOfEdits(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-pass-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, variant + ".xlsx");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27"),
                new SetValueOp("Sheet1", "E5", "boolean", "true")]);
            var assertions = new[]
            {
                Assertion("Sheet1!B1", new { value = 27.0 }),
                Assertion("Sheet1!A1", new { value = "hello" }),
                Assertion("Sheet1!D3", new { value = "old" }),
                Assertion("Sheet1!E5", new { value = true }),
                Assertion("Sheet1!C1", new { formula = "=1+1" }),
                Assertion("Sheet1!F9", new { value = (object?)null })
            }.Select(item => item.Normalize()).ToArray();
            var output = Path.Combine(directory, "passed.xlsx");
            var result = JsonSerializer.SerializeToElement(sessions.Save(id, output, assertions));
            Assert.Contains(result.GetProperty("gates").EnumerateArray(), gate => gate.GetString() == "G7");
            Assert.Equal(assertions.Length, result.GetProperty("assertions").GetArrayLength());
            Assert.All(result.GetProperty("assertions").EnumerateArray(), assertion =>
                Assert.Equal("verified", assertion.GetProperty("status").GetString()));
            Assert.True(File.Exists(output));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void G7ChecksFormulaCachesByTypeAndBlocksWrongAssertions(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, variant + ".xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [
                new SetValueOp("Sheet1", "E5", "formula", "1+1", Operation: "set_formula", ExplicitCache: new FormulaCache("n", "2.50")),
                new SetValueOp("Sheet1", "F5", "formula", "1+1", Operation: "set_formula", ExplicitCache: new FormulaCache("b", "1")),
                new SetValueOp("Sheet1", "G5", "formula", "1+1", Operation: "set_formula", ExplicitCache: new FormulaCache("str", "2")),
                new SetValueOp("Sheet1", "H5", "formula", "1+1", Operation: "set_formula", ExplicitCache: new FormulaCache("e", "#N/A")),
                new SetValueOp("Sheet1", "I5", "formula", "1+1", Operation: "set_formula")]);
            var assertions = new[]
            {
                Assertion("Sheet1!C1", new { value = 2, formula = "1+1" }),
                Assertion("Sheet1!E5", new { value = 2.5, formula = "1+1" }),
                Assertion("Sheet1!F5", new { value = true }),
                Assertion("Sheet1!G5", new { value = "2" }),
                Assertion("Sheet1!H5", new { value = new { error = "#N/A" } }),
                Assertion("Sheet1!I5", new { value = (object?)null })
            }.Select(item => item.Normalize()).ToArray();
            var output = Path.Combine(directory, "typed.xlsx");
            Assert.Equal("verified", JsonSerializer.SerializeToElement(sessions.Save(id, output, assertions))
                .GetProperty("status").GetString());
            var cacheReadback = JsonSerializer.SerializeToElement(P2aGates.ReadCells(output, "Sheet1", ["E5"]))[0];
            Assert.False(cacheReadback.TryGetProperty("CacheType", out _));
            Assert.False(cacheReadback.TryGetProperty("CacheRawValue", out _));
            foreach (var (address, wrong) in new[]
            {
                ("E5", (object)"2.50"), ("F5", 1), ("G5", 2), ("H5", "#N/A"), ("I5", 0), ("C1", 3)
            })
                Assert.Contains(G7Assertions.Check(output, [Assertion("Sheet1!" + address, new { value = wrong }).Normalize()]),
                    issue => issue.Code == "ASSERT_VALUE_MISMATCH");
            var blockedPath = Path.Combine(directory, "blocked.xlsx");
            var blocked = Assert.Throws<SaveBlockedException>(() => sessions.Save(id, blockedPath,
                [Assertion("Sheet1!G5", new { value = 2 }).Normalize()]));
            Assert.Contains(blocked.Issues, issue => issue.Code == "ASSERT_VALUE_MISMATCH");
            Assert.False(File.Exists(blockedPath));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("b", "9", "ASSERT_VALUE_MISMATCH")]
    [InlineData("s", "0", "ASSERT_CACHE_UNSUPPORTED")]
    [InlineData("n", "not-a-number", "ASSERT_VALUE_MISMATCH")]
    public void G7RejectsMalformedOrUnsupportedFormulaCache(string type, string cached, string code)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-invalid-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var output = Path.Combine(directory, "invalid.xlsx");
            using (var store = new PackageStore(Path.Combine(directory, "default.xlsx")))
            {
                var part = store.SheetPart("Sheet1");
                var document = PackageStore.Parse(store.Read(part));
                var cell = Assert.Single(document.GetElementsByTagName("c", PackageStore.Main)
                    .OfType<System.Xml.XmlElement>(), element => element.GetAttribute("r") == "C1");
                cell.SetAttribute("t", type);
                Assert.Single(cell.GetElementsByTagName("v", PackageStore.Main)
                    .OfType<System.Xml.XmlElement>()).InnerText = cached;
                store.Set(part, Encoding.UTF8.GetBytes(document.OuterXml));
                store.Save(output);
            }
            var assertion = Assertion("Sheet1!C1", new { value = false }).Normalize();
            Assert.Contains(G7Assertions.Check(output, [assertion]), issue => issue.Code == code);
            Assert.False(G7Assertions.Matches(Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["C1"])), assertion));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G7BlocksWrongValueOrFormulaAndLeavesNoDestination()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-block-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(Path.Combine(directory, "default.xlsx")))
                .GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            var output = Path.Combine(directory, "blocked.xlsx");
            var assertions = new[]
            {
                Assertion("Sheet1!B1", new { value = 28 }),
                Assertion("Sheet1!C1", new { formula = "9+9" }),
                Assertion("Wrong!A1", new { value = 1 })
            }.Select(item => item.Normalize()).ToArray();
            var blocked = Assert.Throws<SaveBlockedException>(() => sessions.Save(id, output, assertions));
            Assert.Contains(blocked.Issues, issue => issue.Gate == "G7" && issue.Code == "ASSERT_VALUE_MISMATCH");
            Assert.Contains(blocked.Issues, issue => issue.Gate == "G7" && issue.Code == "ASSERT_FORMULA_MISMATCH");
            Assert.Contains(blocked.Issues, issue => issue.Gate == "G7" && issue.Code == "ASSERT_SHEET_MISSING");
            Assert.False(File.Exists(output));
            Assert.Empty(Directory.GetFiles(directory, "*.staging"));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("bom-crlf-standalone")]
    [InlineData("opc-percent-case")]
    [InlineData("new-shared-strings")]
    [InlineData("nested-workbook")]
    public void G7UnchangedChecksExistingAndMissingCellsAndBlocksEdits(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-unchanged-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, variant + ".xlsx");
            foreach (var (target, operation) in new[]
            {
                ("B1", new SetValueOp("Sheet1", "B1", "number", "27")),
                ("F9", new SetValueOp("Sheet1", "F9", "number", "27")),
                ("D3", new SetValueOp("Sheet1", "D3", "blank", null, Operation: "clear", RemoveCell: true))
            })
            {
                using var sessions = new ExcelSessions();
                var id = JsonSerializer.SerializeToElement(sessions.Open(source)).GetProperty("session").GetString()!;
                sessions.Apply(id, 0, [operation]);
                var assertions = new[] { "A1", "B1", "C1", "D3", "F9" }
                    .Select(address => Unchanged("Sheet1!" + address)).ToArray();
                var output = Path.Combine(directory, target + ".xlsx");
                var blocked = Assert.Throws<SaveBlockedException>(() => sessions.Save(id, output, assertions));
                Assert.Equal(["Sheet1!" + target], blocked.Issues.Select(issue => issue.Detail).ToArray());
                Assert.All(blocked.Issues, issue => Assert.Equal("ASSERT_CHANGED", issue.Code));
                Assert.False(File.Exists(output));
                Assert.Empty(Directory.GetFiles(directory, "*.staging"));

                Assert.Equal("verified", JsonSerializer.SerializeToElement(sessions.Save(id, output,
                    assertions.Where(assertion => assertion.Address != target).ToArray()))
                    .GetProperty("status").GetString());
                sessions.Close(id, true);
            }
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void G7UnchangedChecksReferencedSharedStringMarkupNotJustValue()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-phonetic-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, "default.xlsx");
            var output = Path.Combine(directory, "phonetic.xlsx");
            using (var store = new PackageStore(source))
            {
                var part = "xl/sharedStrings.xml";
                var xml = Encoding.UTF8.GetString(store.Read(part));
                Assert.Contains("phoneticPr fontId=\"0\"", xml);
                store.Set(part, Encoding.UTF8.GetBytes(xml.Replace("phoneticPr fontId=\"0\"",
                    "phoneticPr fontId=\"1\"", StringComparison.Ordinal)));
                store.Save(output);
            }
            Assert.Equal("hello", Assert.Single(P2aGates.ReadCells(output, "Sheet1", ["A1"])).Value);
            Assert.Contains(G7Assertions.Check(output, [Unchanged("Sheet1!A1")], source),
                issue => issue.Code == "ASSERT_CHANGED");
            Assert.Empty(G7Assertions.Check(output, [Unchanged("Sheet1!B1")], source));
            Assert.Contains(G7Assertions.Check(output, [Unchanged("Sheet1!A1")]),
                issue => issue.Code == "ASSERT_SOURCE_REQUIRED");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("default")]
    [InlineData("prefixed-x")]
    [InlineData("new-shared-strings")]
    public void G7RichAssertionsReadRunsAndBlockLostFormatting(string variant)
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-rich-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var source = Path.Combine(directory, variant + ".xlsx");
            var richInput = Path.Combine(directory, "rich-input.xlsx");
            RichFixture.Create(source, richInput, variant == "new-shared-strings");
            var correct = Assertion("Sheet1!A1", new { value = "hello bold",
                rich = "<r>hello</r><r b color=\"FF0000\"> bold</r>" }).Normalize();
            var wrong = Assertion("Sheet1!A1", new { value = "hello bold",
                rich = "<r>hello</r><r b=\"0\" color=\"FF0000\"> bold</r>" }).Normalize();
            Assert.Empty(G7Assertions.Check(richInput, [correct]));
            Assert.Contains(G7Assertions.Check(richInput, [wrong]), issue => issue.Code == "ASSERT_RICH_MISMATCH");
            var corrupted = Path.Combine(directory, "corrupted-rich.xlsx");
            using (var store = new PackageStore(richInput))
            {
                var part = variant == "new-shared-strings" ? store.SheetPart("Sheet1") : "xl/sharedStrings.xml";
                var xml = Encoding.UTF8.GetString(store.Read(part));
                Assert.Contains("b val=\"1\"", xml);
                store.Set(part, Encoding.UTF8.GetBytes(xml.Replace("b val=\"1\"", "b val=\"0\"", StringComparison.Ordinal)));
                store.Save(corrupted);
            }
            Assert.Equal("hello bold", Assert.Single(P2aGates.ReadCells(corrupted, "Sheet1", ["A1"])).Value);
            Assert.Contains(G7Assertions.Check(corrupted, [correct]), issue => issue.Code == "ASSERT_RICH_MISMATCH");
            var unmodeled = Path.Combine(directory, "unmodeled-rich.xlsx");
            using (var store = new PackageStore(richInput))
            {
                var part = variant == "new-shared-strings" ? store.SheetPart("Sheet1") : "xl/sharedStrings.xml";
                var xml = Encoding.UTF8.GetString(store.Read(part));
                store.Set(part, Encoding.UTF8.GetBytes(xml.Replace("b val=\"1\"",
                    "b val=\"1\" hidden=\"1\"", StringComparison.Ordinal)));
                store.Save(unmodeled);
            }
            Assert.Contains(G7Assertions.Check(unmodeled, [correct]), issue => issue.Code == "ASSERT_RICH_MISMATCH");
            using var sessions = new ExcelSessions();
            var id = JsonSerializer.SerializeToElement(sessions.Open(richInput)).GetProperty("session").GetString()!;
            sessions.Apply(id, 0, [new SetValueOp("Sheet1", "B1", "number", "27")]);
            var blocked = Path.Combine(directory, "blocked.xlsx");
            Assert.Contains(Assert.Throws<SaveBlockedException>(() => sessions.Save(id, blocked, [wrong])).Issues,
                issue => issue.Code == "ASSERT_RICH_MISMATCH");
            Assert.False(File.Exists(blocked));
            var output = Path.Combine(directory, "verified.xlsx");
            sessions.Save(id, output, [correct]);
            Assert.True(File.Exists(output));
            sessions.Close(id, true);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("<r unknown=\"1\">hello</r>")]
    [InlineData("<r><b/></r>")]
    [InlineData("<r b=\"maybe\">hello</r>")]
    [InlineData("<r color=\"theme:5,2\">hello</r>")]
    [InlineData("<r>hello</i>")]
    [InlineData("<!DOCTYPE rich [<!ENTITY x SYSTEM \"file:///etc/passwd\">]>&x;")]
    public void RichAssertionRejectsMalformedOrUnmodeledMarkup(string rich)
    {
        var error = Record.Exception(() => Assertion("Sheet1!A1", new { rich }).Normalize());
        Assert.True(error is ArgumentException or NotSupportedException, error?.ToString());
    }

    [Fact]
    public void RichAssertionValidatesSupportedRunAttributes()
    {
        var rich = "plain<r b i u sz=\"14\" font=\"Times New Roman\" color=\"theme:5,-0.25\" va=\"superscript\" family=\"2\" charset=\"1\" scheme=\"minor\">run</r>";
        Assert.Equal(rich, Assertion("Sheet1!A1", new { rich }).Normalize().Rich);
        Assert.Throws<NotSupportedException>(() => Assertion("Sheet1!A1", new { rich = new string('x', 8193) }).Normalize());
    }

    [Theory]
    [InlineData("{\"target\":\"Sheet1!A1\",\"unchanged\":false}")]
    [InlineData("{\"target\":\"Sheet1!A1\",\"unchanged\":null}")]
    [InlineData("{\"target\":\"Sheet1!A1\",\"unchanged\":\"true\"}")]
    [InlineData("{\"target\":\"Sheet1!A1\",\"unchanged\":true,\"except\":[\"value\"]}")]
    [InlineData("{\"target\":\"Sheet1!A1\",\"unchanged\":true,\"equals\":{\"value\":\"hello\"}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{\"display\":\"27\"}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{\"value\":27},\"unchanged\":true}")]
    [InlineData("{\"target\":\"B1\",\"equals\":{\"value\":27}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{\"value\":{\"date\":\"2026-10-15\"}}}")]
    [InlineData("{\"target\":\"Sheet1!B1\",\"equals\":{\"value\":1e2147483648}}")]
    public void RejectsUnsupportedOrAmbiguousAssertions(string json)
    {
        var assertion = JsonSerializer.Deserialize<SaveAssertionRequest>(json)!;
        var error = Record.Exception(() => assertion.Normalize());
        Assert.True(error is ArgumentException or NotSupportedException, error?.ToString());
    }

    [Fact]
    public void G7IndependentReaderDetectsCellMismatch()
    {
        var directory = Path.Combine(Path.GetTempPath(), "docloupe-g7-reader-" + Guid.NewGuid().ToString("N"));
        try
        {
            SyntheticFixtures.Create(directory);
            var path = Path.Combine(directory, "default.xlsx");
            Assert.Empty(G7Assertions.Check(path, [Assertion("Sheet1!B1", new { value = 42 }).Normalize()]));
            Assert.Contains(G7Assertions.Check(path, [Assertion("Sheet1!B1", new { value = 43 }).Normalize()]),
                issue => issue.Code == "ASSERT_VALUE_MISMATCH");
            Assert.Empty(G7Assertions.Check(path, [Assertion("Sheet1!B1", new { value = 4.2e1 }).Normalize()]));
            using (var store = new PackageStore(path))
            {
                SetValueEngine.Apply(store, [new SetValueOp("Sheet1", "B1", "number", "0")]);
                var zero = Path.Combine(directory, "zero.xlsx");
                store.Save(zero);
                Assert.Contains(G7Assertions.Check(zero, [Assertion("Sheet1!B1", new { value = 1e-100 }).Normalize()]),
                    issue => issue.Code == "ASSERT_VALUE_MISMATCH");
            }
            Assert.Empty(G7Assertions.Check(path, [Assertion("Sheet1!C1", new { value = 2 }).Normalize()]));
            Assert.Contains(G7Assertions.Check(path, [Assertion("Sheet1!C1", new { value = 3 }).Normalize()]),
                issue => issue.Code == "ASSERT_VALUE_MISMATCH");
        }
        finally { Directory.Delete(directory, true); }
    }

    private static SaveAssertionRequest Assertion(string target, object equals) => new()
    {
        Target = target, Expected = JsonSerializer.SerializeToElement(equals)
    };

    private static ValueAssertion Unchanged(string target) => JsonSerializer.Deserialize<SaveAssertionRequest>(
        JsonSerializer.Serialize(new { target, unchanged = true }))!.Normalize();
}
