using System.IO.Compression;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

if (args is not [var serverDll]) throw new ArgumentException("Usage: <server-dll-path>");
var directory = Path.Combine(Path.GetTempPath(), "docloupe-s3-" + Guid.NewGuid().ToString("N"));
try
{
    var source = SyntheticFixtures.Create(directory)[0];
    var transport = new StdioClientTransport(new StdioClientTransportOptions
    {
        Command = "dotnet",
        Arguments = [Path.GetFullPath(serverDll)],
        Name = "p2a-s3"
    });
    await using var client = await McpClient.CreateAsync(transport);
    var tools = await client.ListToolsAsync();
    Console.WriteLine("tools=" + string.Join(',', tools.Select(tool => tool.Name)));
    if (!tools.Any(tool => tool.Name == "excel_undo")) throw new InvalidOperationException("Undo tool is missing");
    if (!tools.Any(tool => tool.Name == "excel_status")) throw new InvalidOperationException("Status tool is missing");
    if (!tools.Any(tool => tool.Name == "excel_peek")) throw new InvalidOperationException("Peek tool is missing");
    if (!tools.Any(tool => tool.Name == "excel_verify")) throw new InvalidOperationException("Verify tool is missing");
    var verified = await client.CallToolAsync("excel_verify", new Dictionary<string, object?> { ["after_path"] = source });
    if (verified.IsError == true || verified.StructuredContent?.GetProperty("data").GetProperty("status").GetString() != "unverified" ||
        verified.StructuredContent?.GetProperty("data").GetProperty("partial").GetBoolean() != true ||
        verified.StructuredContent?.GetProperty("data").GetProperty("unverified_gates").GetArrayLength() == 0 ||
        verified.StructuredContent?.GetProperty("warnings").GetArrayLength() == 0)
        throw new InvalidOperationException("Read-only verification claimed full success: " + verified.StructuredContent?.GetRawText());
    if (verified.StructuredContent?.GetProperty("data").GetProperty("schema_issues").GetArrayLength() != 0)
        throw new InvalidOperationException("Synthetic fixture must be schema-valid: " + verified.StructuredContent?.GetRawText());
    var compared = await client.CallToolAsync("excel_verify", new Dictionary<string, object?>
    {
        ["after_path"] = source, ["before_path"] = source
    });
    if (compared.IsError == true || compared.StructuredContent?.GetProperty("data").GetProperty("status").GetString() != "unverified" ||
        compared.StructuredContent?.GetProperty("data").GetProperty("differences").GetArrayLength() != 0)
        throw new InvalidOperationException("No-change comparison failed: " + compared.StructuredContent?.GetRawText());
    var lostEdit = await client.CallToolAsync("excel_verify", new Dictionary<string, object?>
    {
        ["after_path"] = source, ["before_path"] = source,
        ["assert"] = new[] { new { target = "Sheet1!B1", equals = new { value = 99 } } }
    });
    if (lostEdit.IsError != true || lostEdit.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "ASSERTION_FAILED" ||
        lostEdit.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("differences").GetArrayLength() != 0 ||
        lostEdit.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("assertion_issues")[0]
            .GetProperty("Code").GetString() != "ASSERT_VALUE_MISMATCH")
        throw new InvalidOperationException("Lost edit was not caught by verification assertion: " + lostEdit.StructuredContent?.GetRawText());
    var changedWorkbook = Path.Combine(directory, "compare-changed.xlsx");
    File.Copy(source, changedWorkbook);
    using (var archive = ZipFile.Open(changedWorkbook, ZipArchiveMode.Update))
    {
        var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
        string xml;
        using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
        entry.Delete();
        using var writer = new StreamWriter(archive.CreateEntry("xl/worksheets/sheet1.xml").Open());
        writer.Write(xml.Replace("<v>42</v>", "<v>43</v>", StringComparison.Ordinal));
    }
    var comparedChange = await client.CallToolAsync("excel_verify", new Dictionary<string, object?>
    {
        ["after_path"] = changedWorkbook, ["before_path"] = source, ["max_differences"] = 1
    });
    if (comparedChange.IsError != true || comparedChange.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "PRESERVATION_FAILED" ||
        comparedChange.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("differences").GetArrayLength() != 1)
        throw new InvalidOperationException("Undeclared change passed comparison: " + comparedChange.StructuredContent?.GetRawText());
    var broken = Path.Combine(directory, "broken.xlsx");
    File.Copy(source, broken);
    using (var archive = ZipFile.Open(broken, ZipArchiveMode.Update))
        archive.GetEntry("_rels/.rels")!.Delete();
    var invalidVerification = await client.CallToolAsync("excel_verify", new Dictionary<string, object?> { ["after_path"] = broken });
    if (invalidVerification.IsError != true || invalidVerification.StructuredContent?.GetProperty("ok").GetBoolean() != false ||
        invalidVerification.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "PACKAGE_INVALID" ||
        invalidVerification.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("status").GetString() != "failed")
        throw new InvalidOperationException("Corrupt workbook was accepted: " + invalidVerification.StructuredContent?.GetRawText());
    var invalidZip = Path.Combine(directory, "invalid-zip.xlsx");
    File.WriteAllText(invalidZip, "not a zip archive");
    var invalidZipResult = await client.CallToolAsync("excel_verify", new Dictionary<string, object?> { ["after_path"] = invalidZip });
    if (invalidZipResult.IsError != true || invalidZipResult.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "PACKAGE_INVALID" ||
        invalidZipResult.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("package_issues")[0]
            .GetProperty("Code").GetString() != "INVALID_PACKAGE")
        throw new InvalidOperationException("Invalid ZIP was not reported as a package error: " + invalidZipResult.StructuredContent?.GetRawText());
    var wrongType = Path.Combine(directory, "wrong-type.xlsx");
    File.Copy(source, wrongType);
    using (var archive = ZipFile.Open(wrongType, ZipArchiveMode.Update))
    {
        var entry = archive.GetEntry("[Content_Types].xml")!;
        string contentTypes;
        using (var reader = new StreamReader(entry.Open())) contentTypes = reader.ReadToEnd();
        entry.Delete();
        using var writer = new StreamWriter(archive.CreateEntry("[Content_Types].xml").Open());
        writer.Write(contentTypes.Replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/xml", StringComparison.Ordinal));
    }
    var wrongTypeResult = await client.CallToolAsync("excel_verify", new Dictionary<string, object?> { ["after_path"] = wrongType });
    if (wrongTypeResult.IsError != true || wrongTypeResult.StructuredContent is null ||
        wrongTypeResult.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "PACKAGE_INVALID" ||
        !wrongTypeResult.StructuredContent.Value.GetProperty("error").GetProperty("details").GetProperty("package_issues")
            .EnumerateArray().Any(issue => issue.GetProperty("Code").GetString() == "WORKBOOK_CONTENT_TYPE_MISMATCH"))
        throw new InvalidOperationException("Wrong workbook content type was accepted: " + wrongTypeResult.StructuredContent?.GetRawText());
    var compressed = Path.Combine(directory, "compressed.xlsx");
    File.Copy(source, compressed);
    using (var archive = ZipFile.Open(compressed, ZipArchiveMode.Update))
    using (var writer = new StreamWriter(archive.CreateEntry("xl/compressed.xml", CompressionLevel.Optimal).Open()))
        writer.Write("<root>" + new string('a', 1_000_001 - 13) + "</root>");
    var compressedResult = await client.CallToolAsync("excel_verify", new Dictionary<string, object?> { ["after_path"] = compressed });
    if (compressedResult.IsError != true || compressedResult.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "PACKAGE_INVALID" ||
        compressedResult.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("package_issues")
            .EnumerateArray().Any(issue => issue.GetProperty("Code").GetString() == "LIMIT_COMPRESSION_RATIO") != true)
        throw new InvalidOperationException("High-compression ZIP was accepted: " + compressedResult.StructuredContent?.GetRawText());
    var peek = await client.CallToolAsync("excel_peek", new Dictionary<string, object?>
    {
        ["path"] = source, ["detail"] = "preview", ["sheet"] = "Sheet1", ["max_rows"] = 3, ["max_cols"] = 4
    });
    if (peek.IsError == true || peek.StructuredContent?.GetProperty("data").GetProperty("preview")[0]
        .GetProperty("markdown").GetString()?.Contains("hello", StringComparison.Ordinal) != true ||
        peek.StructuredContent?.GetProperty("data").GetProperty("sheets")[0]
            .GetProperty("used_range").GetString() != "A1:D3" ||
        peek.StructuredContent?.GetProperty("data").GetProperty("used_range_basis").GetString() != "explicit_cells")
        throw new InvalidOperationException("Sessionless peek failed: " + peek.StructuredContent?.GetRawText());
    var emptyStatus = await client.CallToolAsync("excel_status", new Dictionary<string, object?>());
    if (emptyStatus.IsError == true || emptyStatus.StructuredContent?.GetProperty("data").GetProperty("sessions").GetArrayLength() != 0)
        throw new InvalidOperationException("Empty status failed: " + emptyStatus.StructuredContent?.GetRawText()
            + " text=" + emptyStatus.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text
            + " schema=" + tools.Single(tool => tool.Name == "excel_status").ProtocolTool.InputSchema.GetRawText());
    var result = await client.CallToolAsync("excel_open", new Dictionary<string, object?> { ["path"] = source });
    Console.WriteLine("structured_content=" + result.StructuredContent?.GetRawText());
    Console.WriteLine("text_content=" + (result.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text ?? "<absent>"));
    if (result.IsError == true || result.StructuredContent is null || result.Content.OfType<TextContentBlock>().Count() != 1)
        throw new InvalidOperationException("Dual-channel MCP tool response not received by SDK client");
    var session = result.StructuredContent.Value.GetProperty("data").GetProperty("session").GetString()!;
    var rangeRead = await client.CallToolAsync("excel_read", new Dictionary<string, object?>
    {
        ["session"] = session, ["sheet"] = "Sheet1", ["target"] = "A1:C2"
    });
    if (rangeRead.IsError == true ||
        rangeRead.StructuredContent?.GetProperty("data").GetProperty("cells").GetArrayLength() != 3)
        throw new InvalidOperationException("Bounded range read failed: " + rangeRead.StructuredContent?.GetRawText());
    var typedRead = await client.CallToolAsync("excel_read", new Dictionary<string, object?>
    {
        ["session"] = session, ["target"] = "Sheet1!A1:C2", ["view"] = "values"
    });
    if (typedRead.IsError == true || typedRead.StructuredContent?.GetProperty("data").GetProperty("rows")[0][0].GetString() != "hello" ||
        typedRead.StructuredContent?.GetProperty("data").GetProperty("rows")[0][1].GetInt32() != 42 ||
        typedRead.StructuredContent?.GetProperty("data").GetProperty("rows")[1][0].ValueKind != System.Text.Json.JsonValueKind.Null)
        throw new InvalidOperationException("Typed values view failed: " + typedRead.StructuredContent?.GetRawText());
    var markdownRead = await client.CallToolAsync("excel_read", new Dictionary<string, object?>
    {
        ["session"] = session, ["target"] = "Sheet1!B1:C2", ["view"] = "markdown"
    });
    if (markdownRead.IsError == true ||
        markdownRead.StructuredContent?.GetProperty("data").GetProperty("markdown").GetString()?.Contains("| 1 | 42 | 2 ƒ =1+1 |", StringComparison.Ordinal) != true)
        throw new InvalidOperationException("Markdown range failed: " + markdownRead.StructuredContent?.GetRawText());
    var found = await client.CallToolAsync("excel_find", new Dictionary<string, object?>
    {
        ["session"] = session, ["query"] = new { text = "HELLO" },
        ["scope"] = new { sheet = "Sheet1", target = "A1:D3" }, ["in"] = "value"
    });
    if (found.IsError == true || found.StructuredContent?.GetProperty("data").GetProperty("partial").GetBoolean() != true ||
        found.StructuredContent?.GetProperty("data").GetProperty("matches")[0].GetProperty("addr").GetString() != "Sheet1!A1")
        throw new InvalidOperationException("Bounded find failed: " + found.StructuredContent?.GetRawText());
    var workbookFind = await client.CallToolAsync("excel_find", new Dictionary<string, object?>
    {
        ["session"] = session, ["query"] = new { text = "HELLO" }
    });
    if (workbookFind.IsError == true ||
        workbookFind.StructuredContent?.GetProperty("data").GetProperty("matches")[0].GetProperty("addr").GetString() != "Sheet1!A1")
        throw new InvalidOperationException("Unscoped bounded find failed: " + workbookFind.StructuredContent?.GetRawText());
    var readSchema = tools.Single(tool => tool.Name == "excel_read").ProtocolTool.InputSchema;
    if (readSchema.GetProperty("properties").GetProperty("target").GetProperty("oneOf").GetArrayLength() != 2)
        throw new InvalidOperationException("Read target array is missing from MCP input schema: " + readSchema.GetRawText());
    var listedRead = await client.CallToolAsync("excel_read", new Dictionary<string, object?>
    {
        ["session"] = session, ["sheet"] = "Sheet1", ["target"] = new[] { "A1", "B1:C2" }
    });
    if (listedRead.IsError == true ||
        listedRead.StructuredContent?.GetProperty("data").GetProperty("cells").GetArrayLength() != 3)
        throw new InvalidOperationException("Read target list failed: " + listedRead.StructuredContent?.GetRawText()
            + " text=" + listedRead.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text
            + " schema=" + tools.Single(tool => tool.Name == "excel_read").ProtocolTool.InputSchema.GetRawText());
    var invalidRead = await client.CallToolAsync("excel_read", new Dictionary<string, object?>
    {
        ["session"] = session, ["sheet"] = "Sheet1", ["target"] = new object[] { "A1", 7 }
    });
    if (invalidRead.IsError != true)
        throw new InvalidOperationException("Non-string read target was accepted");
    var stale = await client.CallToolAsync("excel_apply", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 0, ["sheet"] = "Sheet1",
        ["ops"] = new[] { new { op = "set_value", target = "B1", value = 99, expect = new { value = 0 } } }
    });
    if (stale.IsError != true || stale.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "PRECONDITION_FAILED" ||
        stale.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("index").GetInt32() != 0)
        throw new InvalidOperationException("MCP precondition failure was not structured: " + stale.StructuredContent?.GetRawText());
    var expandedFailure = await client.CallToolAsync("excel_apply", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 0, ["sheet"] = "Sheet1",
        ["ops"] = new object[]
        {
            new { op = "set_value", target = "F5:G5", value = 6 },
            new { op = "clear", target = "B1", expect = new { value = 0 } }
        }
    });
    if (expandedFailure.IsError != true ||
        expandedFailure.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "PRECONDITION_FAILED" ||
        expandedFailure.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("index").GetInt32() != 1)
        throw new InvalidOperationException("Expanded batch reported a cell index instead of the request index: "
            + expandedFailure.StructuredContent?.GetRawText());
    var planned = await client.CallToolAsync("excel_apply", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 0, ["sheet"] = "Sheet1", ["dry_run"] = true,
        ["ops"] = new[] { new { op = "set_value", target = "B1", value = 99, expect = new { value = 42 } } }
    });
    if (planned.IsError == true || planned.StructuredContent?.GetProperty("data").GetProperty("dry_run").GetBoolean() != true ||
        planned.StructuredContent?.GetProperty("data").GetProperty("revision_after").GetInt32() != 0 ||
        planned.StructuredContent?.GetProperty("data").GetProperty("intent").GetArrayLength() != 1)
        throw new InvalidOperationException("MCP dry run failed: " + planned.StructuredContent?.GetRawText());
    var applied = await client.CallToolAsync("excel_apply", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 0, ["sheet"] = "Sheet1",
        ["ops"] = new object[] { new { op = "set_value", sheet = "Sheet1", target = "B1", value = 99, expect = new { value = 42 } },
            new { op = "set_formula", sheet = "Sheet1", target = "C1", formula = "=B1+2", cache = "keep", expect = new { formula = "=1+1", value = 2 } },
            new { op = "set_values", sheet = "Sheet1", target = "D4", values = new object?[][] { [4, "batch"] } },
            new { op = "set_value", sheet = "Sheet1", target = "F5:G5", value = 6 },
            new { op = "fill", sheet = "Sheet1", target = "H6:I6", value = 8 },
            new { op = "fill", sheet = "Sheet1", target = "J7:K7", series = new { start = -3, step = 2 } },
            new { op = "clear", sheet = "Sheet1", target = "A1", what = new[] { "values" }, remove_cells = false },
            new { op = "clear", sheet = "Sheet1", target = "L8" },
            new { op = "clear", sheet = "Sheet1", target = "D3", remove_cells = true },
            new { op = "set_value", sheet = "Sheet1", target = "M9", value = new { error = "#N/A" } },
            new { op = "fill", sheet = "Sheet1", target = "N10:O10", series = new { start = 0.1, step = 0.2 } },
            new { op = "set_formula", sheet = "Sheet1", target = "P11", formula = "2+3", cache = (object)new { value = 5 } },
            new { op = "set_formula", sheet = "Sheet1", target = "Q11", formula = "1/0", cache = (object)new { value = new { error = "#DIV/0!" } } } }
    });
    if (applied.IsError == true) throw new InvalidOperationException("Apply failed: " + applied.StructuredContent?.GetRawText());
    var status = await client.CallToolAsync("excel_status", new Dictionary<string, object?> { ["session"] = session });
    if (status.IsError == true || status.StructuredContent?.GetProperty("data").GetProperty("revision").GetInt32() != 1 ||
        status.StructuredContent?.GetProperty("data").GetProperty("ledger").GetArrayLength() != 1)
        throw new InvalidOperationException("Session status failed: " + status.StructuredContent?.GetRawText());
    var staleFormula = await client.CallToolAsync("excel_apply", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 1, ["sheet"] = "Sheet1",
        ["ops"] = new[] { new { op = "clear", target = "C1", expect = new { formula = "1+1" } } }
    });
    if (staleFormula.IsError != true || staleFormula.StructuredContent?.GetProperty("error").GetProperty("code").GetString() != "PRECONDITION_FAILED" ||
        staleFormula.StructuredContent?.GetProperty("error").GetProperty("details").GetProperty("actual").GetProperty("Formula").GetString() != "B1+2")
        throw new InvalidOperationException("MCP formula precondition did not catch a stale revision: " + staleFormula.StructuredContent?.GetRawText());
    var invalidOutput = Path.Combine(directory, "invalid.xlsx");
    var rejected = await client.CallToolAsync("excel_save", new Dictionary<string, object?>
    {
        ["session"] = session, ["mode"] = "copy", ["path"] = invalidOutput,
        ["assert"] = new[] { new { target = "Sheet1!B1", equals = new { display = "99" } } }
    });
    if (rejected.IsError != true || File.Exists(invalidOutput))
        throw new InvalidOperationException("Unsupported assertion was not rejected");
    var changedOutput = Path.Combine(directory, "changed.xlsx");
    var changed = await client.CallToolAsync("excel_save", new Dictionary<string, object?>
    {
        ["session"] = session, ["mode"] = "copy", ["path"] = changedOutput,
        ["assert"] = new[] { new { target = "Sheet1!B1", unchanged = true } }
    });
    if (changed.IsError != true || File.Exists(changedOutput))
        throw new InvalidOperationException("Changed cell passed an unchanged assertion");
    var output = Path.Combine(directory, "written.xlsx");
    var saved = await client.CallToolAsync("excel_save", new Dictionary<string, object?>
    {
        ["session"] = session, ["mode"] = "copy", ["path"] = output,
        ["assert"] = new object[] { new { target = "Sheet1!B1", equals = new { value = 99 } },
            new { target = "Sheet1!C1", equals = new { formula = "B1+2", value = 2 } },
            new { target = "Sheet1!D4", equals = new { value = 4 } },
            new { target = "Sheet1!E4", equals = new { value = "batch" } },
            new { target = "Sheet1!F5", equals = new { value = 6 } },
            new { target = "Sheet1!G5", equals = new { value = 6 } },
            new { target = "Sheet1!H6", equals = new { value = 8 } },
            new { target = "Sheet1!I6", equals = new { value = 8 } },
            new { target = "Sheet1!J7", equals = new { value = -3 } },
            new { target = "Sheet1!K7", equals = new { value = -1 } },
            new { target = "Sheet1!A1", equals = new { value = System.Text.Json.JsonSerializer.SerializeToElement<object?>(null) } },
            new { target = "Sheet1!L8", equals = new { value = System.Text.Json.JsonSerializer.SerializeToElement<object?>(null) } },
            new { target = "Sheet1!D3", equals = new { value = System.Text.Json.JsonSerializer.SerializeToElement<object?>(null) } },
            new { target = "Sheet1!M9", equals = new { value = new { error = "#N/A" } } },
            new { target = "Sheet1!N10", equals = new { value = 0.1 } },
            new { target = "Sheet1!O10", equals = new { value = 0.3 } },
            new { target = "Sheet1!P11", equals = new { formula = "2+3", value = 5 } },
            new { target = "Sheet1!Q11", equals = new { formula = "1/0", value = (object)new { error = "#DIV/0!" } } },
            new { target = "Sheet1!B2", unchanged = true } }
    });
    if (saved.IsError == true || saved.StructuredContent?.GetProperty("data").GetProperty("status").GetString() != "verified")
        throw new InvalidOperationException("Verified save failed: " + saved.StructuredContent?.GetRawText());
    var cachedFormula = await client.CallToolAsync("excel_read", new Dictionary<string, object?>
    {
        ["session"] = session, ["sheet"] = "Sheet1", ["target"] = "C1"
    });
    if (cachedFormula.IsError == true || cachedFormula.StructuredContent?.GetProperty("data")
            .GetProperty("cells")[0].GetProperty("Value").GetString() != "2")
        throw new InvalidOperationException("Formula cache was not retained");
    var savedSession = await client.CallToolAsync("excel_open", new Dictionary<string, object?> { ["path"] = output });
    if (savedSession.IsError == true) throw new InvalidOperationException("Saved file could not be reopened");
    var savedId = savedSession.StructuredContent!.Value.GetProperty("data").GetProperty("session").GetString()!;
    foreach (var (address, expected) in new[] { ("P11", "5"), ("Q11", "#DIV/0!") })
    {
        var read = await client.CallToolAsync("excel_read", new Dictionary<string, object?>
        {
            ["session"] = savedId, ["sheet"] = "Sheet1", ["target"] = address
        });
        if (read.IsError == true || read.StructuredContent?.GetProperty("data")
                .GetProperty("cells")[0].GetProperty("Value").GetString() != expected)
            throw new InvalidOperationException("Typed formula cache was not retained at " + address);
    }
    var savedClosed = await client.CallToolAsync("excel_close", new Dictionary<string, object?>
    {
        ["session"] = savedId, ["discard_unsaved"] = false
    });
    if (savedClosed.IsError == true) throw new InvalidOperationException("Saved session close failed");
    var verifiedGates = saved.StructuredContent.Value.GetProperty("data").GetProperty("gates").EnumerateArray()
        .Select(gate => gate.GetString()).ToArray();
    if (!verifiedGates.Contains("G6") || !verifiedGates.Contains("G7"))
        throw new InvalidOperationException("G6 or G7 was not checked");
    Console.WriteLine("saved_status=" + saved.StructuredContent.Value.GetProperty("data").GetProperty("status").GetString());
    var undone = await client.CallToolAsync("excel_undo", new Dictionary<string, object?>
    {
        ["session"] = session, ["base_revision"] = 1, ["to_revision"] = 0
    });
    if (undone.IsError == true || undone.StructuredContent?.GetProperty("data").GetProperty("revision").GetInt32() != 0)
        throw new InvalidOperationException("Undo failed: " + undone.StructuredContent?.GetRawText());
    var closed = await client.CallToolAsync("excel_close", new Dictionary<string, object?>
    {
        ["session"] = session
    });
    if (closed.IsError == true) throw new InvalidOperationException("Close failed");
}
finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
