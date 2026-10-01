# S2b and P1 read/verify foundation — 2026-10-01

## S2b result

The spike meets the three **narrow, measured** questions for the tested single-cell edit. It is not a general-purpose OOXML writer and does not complete P1 or P2a.

1. `RawWorksheetEditor` reads every `.xml` and `.rels` ZIP entry directly into `XmlDocument` with external resolution disabled and DTDs prohibited, including the original namespace-prefixed `[Content_Types].xml`. It constructs a detached Open XML SDK `Worksheet` from the sheet children (`InnerXml`) and changes its first cell. The same edit is then **re-applied by hand** to the original `XmlDocument`; the typed SDK DOM only cross-checks the cell reference and does not produce the output. S2b therefore demonstrates *System.Xml DOM + lexical splice*, not lossless writing of typed SDK edits (see review finding 6). No `SpreadsheetDocument.Open` or `System.IO.Packaging` API is used. Binary parts are copied without parsing.
2. The XML DOM writer preserves root/element prefixes and namespace declarations. A quote-aware XML token scanner locates the first edited cell in both the original XML and a DOM-serialized version; it splices **only that cell's serialized XML** into the original UTF-8 bytes. The XML declaration, line endings and every byte outside the edited cell consequently remain identical. All 8 external sources passed:
   - the prefix/suffix byte comparison. This holds by construction, because it uses the same scanner as the splice;
   - an independent G3;
   - edited-cell readback.

   During review, an independent re-check used Python's expat to locate the first main-namespace `c` element. It confirmed on 8/8 that every differing byte lies inside that cell, and that every other part is identical. Their seven manifest hashes still match; the eighth (`00`) was checked before and after the probe. ZIP container bytes and lexical details *inside* the edited cell are not promised to be identical.
3. A synthetic `x:`-prefixed worksheet with a non-default XML declaration, `mc:Ignorable`, `x14ac:dyDescent` and an unknown `extLst` passed. A separately compiled **SDK-free** verifier (`src/DocLoupe.Excel.Verify`) rejects EX-04 (`UNDECLARED_MC_PREFIX`) and an EX-03 namespace-stripping mutation (`ATTRIBUTE_NAMESPACE_CHANGED`). Its comparison excludes only the *children* of the specifically edited cell; namespace and declarations on the remaining nodes are checked. Ten xUnit cases pass: seven original, plus three regression tests added in review.

All eight real fixtures: `source=same`, `untouched=same`, `outside_cell=same`, `prolog=same`, `prefixes=same`, `declarations=same`, `g3=pass`, `edit=ok`. The first whole-document `XmlDocument` serialization **did not** meet this stronger lexical requirement on 06/07: it normalized CRLF to LF outside the edited cell. The targeted lexical splice removed that regression. The generic XML DOM parsed 10–28 XML/relationship parts per workbook; detached typed worksheet loading succeeded for every first sheet, including all three `ns0:Types` packages. Output workbooks and synthetic inputs remain ignored under `spikes/S2b/out/`.

### Excel native open check

An isolated Excel COM instance opened the **original source and edited output** with `ReadOnly=true`, `UpdateLinks=0`, `AutomationSecurity=ForceDisable`, `DisplayAlerts=false`, and no save:

| Fixture | Original | S2b output | Source hash |
|---|---|---|---|
| `02-table-metadata-source.xlsx` | refused (`800A03EC`) | refused (`800A03EC`) | unchanged |
| `05-advanced-package-source.xlsm` | refused (`800A03EC`) | refused (`800A03EC`) | unchanged |
| `07-real-package-source.xlsx` | opened read-only | opened read-only | unchanged |

Changing **only** the content-types root from `<ns0:Types>` to a default-namespace `<Types>` in throwaway copies did **not** change any of these three outcomes. Therefore the two refusals must not be attributed to `ns0:` alone. The connected spreadsheet validator marked all three edited packages `valid` (and reported detached advanced-part warnings for 05/07), which is not equivalent to Excel opening them. The COM check does **not** establish absence of a repair prompt when `DisplayAlerts=false`; a separate S5 open/repair oracle remains necessary. No Excel process remained after the check.

### S2b boundary before P2a

- Proven: direct per-XML-part DOM loading, a single first-cell edit with the bytes outside that cell unchanged, independent G3 on the tested cases, no changes to other part contents.
- Not proven: UTF-16 sheets (the writer fails closed), arbitrary cell/structural/rich-text operations, all markup-compatibility constructs and actual Excel repair detection. The currently tested splice targets only the first cell; generalizing DOM-to-XML transfer and coordinating multiple edits is P2a work, not a spike result.
- Treat 02/05 as **native-Excel-refused at baseline**, not as success or a new regression. Before requiring Excel-open success for them, obtain a known-openable equivalent fixture or determine their independent defects. `07` is the `ns0:` positive native-open control.

## P1 work in parallel

`src/DocLoupe.Excel.Verify` is a read-only, SDK-free foundation using `ZipArchive` and `XmlReader`. `WorkbookReader.Peek` resolves workbook relationships and streams the first worksheet's bounded cell preview. `VerifyPartial` checks required top-level entries, duplicate names, content-types root, XML well-formedness and G3 across XML/relationship parts; it returns **`unverified`**, never `verified`, when these checks pass. It returns `failed` for EX-04. The cell preview exposes **raw**, not display-resolved or shared-string-resolved, values.

This is **not P1 completion**. Remaining: full G1 graph/content-type/ZIP limits, G2 schema, complete G5 preservation with independent snapshots, G6 advanced-part comparison, paging and other read/find/inspect tools, MCP response envelopes, corruption/equivalence corpus and cross-platform CI. G4 requires an intent/session and cannot be inferred from a two-file comparison. The eight source files are readable by this partial P1 reader; all correctly report `unverified` with the missing gates listed. Ten self-contained tests pass.

## Review findings (2026-10-01)

**Fixed in `src/DocLoupe.Excel.Verify`**, each with a regression test. All three were reproduced against the original code before fixing.

1. **Phonetic text in the cell value.** `WorkbookReader.ReadCells` appended every `<t>`, including the phonetic guide under `<rPh>`: a cell `漢字` with guide `かんじ` read as `漢字かんじ` (fixture 04 has phonetics). `rPh` and `phoneticPr` are now skipped.
2. **Cached value after a formula was lost.** `ReadElementContentAsString()` already advances to the next node, and the loop then called `Read()` again, which skipped the `<v>` that directly follows `<f>`. `<f>A1*2</f><v>20</v>` read as formula `A1*2` with no value. The loop now advances only when nothing consumed the node.
3. **G3 missed `mc:Choice/@Requires`.** This attribute is unqualified, but its tokens are prefixes; an undeclared one passed. It is now checked like `mc:Ignorable`, and its value is part of the compared shape.

**Requirements carried into P2a:**

4. **Make the outside-the-edit check independent.** The spike's `outside_cell` check reuses the splice scanner, so it cannot disagree with the splice. P2a needs an independent check in `Verify`, e.g. an XmlReader-located span compared against a raw byte diff.
5. **Make the splice locator namespace-aware.** It currently matches the qualified-name string (`c` / `x:c`) and does not cross-check the raw span's `r` against the DOM cell. That is acceptable for the first cell, but it is the same risk class as the legacy regex. The P2a version must:
   - locate the element namespace-aware;
   - verify `r` and fail closed on any disagreement;
   - have unit tests for comments and CDATA inside a cell, `>` inside attribute values, self-closing cells, prefixed names, a BOM, CRLF, and same-named elements in other namespaces.
6. **Decide the role of Open XML SDK at the start of P2a.** The S2b output is produced by System.Xml, not by the SDK (see S2b result 1). README decision D2 ("Open XML SDK for direct DOM edits") must be re-decided. One candidate: a System.Xml-based writer core, with the SDK used for G2 schema validation and typed fragment construction.
7. **Broaden edit coverage.** In six of eight fixtures the first cell is already an inline string (a 4-byte difference), so no conversion from shared-string, number or formula cells was exercised. The P2a matrix must cover:
   - arbitrary target cells;
   - several edits in one sheet;
   - cell-type changes.

**Notes:**

8. **P1 G1 is incomplete.** It hard-codes `xl/workbook.xml` and an `xl` base for relative targets, looks entries up case-sensitively, and does not percent-decode. Before G1 can report `verified`, it must resolve the main part through `_rels/.rels` and follow OPC part-name rules (relative to the source part, case-insensitive).
9. **Fixtures 02 and 05 are refused by Excel at baseline.** Record this alongside the local fixture manifest, so that no test expects Excel to open them.

## Reproduce

From `servers/excel_cs` on Windows:

```powershell
dotnet run --project spikes/S2b/S2b.csproj -- probe 'D:\data-test\excel-preservation-fixtures\sources' 'spikes/S2b/out'
dotnet run --project spikes/S2b/S2b.csproj -- p1 'D:\data-test\excel-preservation-fixtures\sources' 'spikes/S2b/out/EX04_invalid.xlsx'
dotnet run --project spikes/S2b/S2b.csproj -- normalize-types 'D:\data-test\excel-preservation-fixtures\sources' 'spikes/S2b/out/normalized'
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File 'spikes/S2b/excel-open.ps1'
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File 'spikes/S2b/excel-open.ps1' -SourceRoot (Resolve-Path 'spikes/S2b/out').Path -Suffix '-s2b'
dotnet test tests/DocLoupe.Excel.Verify.Tests/DocLoupe.Excel.Verify.Tests.csproj
```

The `normalize-types` command creates derived files in the ignored output directory; pass `-SourceRoot (Resolve-Path 'spikes/S2b/out/normalized').Path` to the same Excel script to repeat that control. No external fixture is copied into tracked files.
