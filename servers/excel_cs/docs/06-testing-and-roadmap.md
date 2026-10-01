# 06 — Testing, spikes and roadmap

## 1. Principle

**The verifier is the product.** If the verifier says `verified` while the file is wrong, the agent is misled, and that is worse than having no tool. So the verifier is tested like a safety system, on two measures:

- **Sensitivity:** every real loss in the corruption corpus is detected by the expected gate. Target: 100%. Any miss is a release blocker.
- **Specificity:** semantically equal files are not flagged. Each normalization rule has bidirectional fixtures, and the false-positive rate is tracked on the equivalence corpus.

The legacy server's 315 Python tests mostly assert on the legacy session model and its outputs. They are not ported. Their scenarios are reused through the parity harness (§2.6).

## 2. Test layers

### 2.1 Unit tests (`Model`, `Package`, `Formula`)

- **A1 parsing and formatting:** property tests `format(parse(a)) == a`; quoted sheet names; limits; whole rows/columns.
- **Rich-text markup:** `parse(render(runs)) == runs` (FsCheck); escapes; explicit false (`b="0"`); whitespace and `\n`.
- **Graphemes:** NFC and NFD Vietnamese (`Việt`, `Đỏ đậm`), combining sequences, emoji ZWJ sequences, CJK with phonetics. `match` under `normalize: "nfc"` and `"none"`. No span may split a cluster.
- **Style patch semantics:** present / `null` / absent, explicit `false`, color forms, xf reuse without renumbering.
- **Display strings:** a table of ≥ 300 (value, numFmt) pairs with the Excel-produced display text under en-US regional settings, plus `date1904` cases and locale-dependent formats. It is captured once on the developer machine via the Excel oracle (Q4: no Office in CI) and committed as data, since we generated it ourselves.
- **PackageStore:** an untouched entry is copied with identical decompressed bytes; `[Content_Types].xml` and `.rels` are edited per entry; zip-bomb limits.

### 2.2 Op tests (`Engine`)

For each op in the catalog (03 §4.2), and each mode (DOM and streaming where relevant):

1. Apply it to a fixture.
2. Assert that the declared effects equal the expected effect list (a golden file).
3. Assert the readback.
4. Assert that the in-session post-apply check passes.
5. Save, then assert that G1–G7 are `verified`.
6. Run the same op with `dry_run` and assert that the plan equals the applied diff.

Negative cases per op cover:

- invalid input;
- precondition failure;
- ambiguous targets;
- merged non-origin cells;
- rich-text policy;
- partial-batch failure, where the session must be unchanged.

### 2.3 Mutation matrix (op × fixture)

This follows audit recommendation 4. Every mutating op runs against every fixture, and each saved copy is checked five ways:

1. the C# gates;
2. `excel_verify` against O, which must show only declared differences;
3. the **legacy Python verifier** as a second opinion, with known legacy false positives (V-02, V-07) filtered by rule id;
4. a LibreOffice open + render smoke test;
5. on the Windows job, Excel `open_check`.

Fixture groups follow the audit: cells, rows/columns, sheets, styles, formulas, validation, tables, drawings, printing, metadata, package edits.

### 2.4 Verifier sensitivity: corruption corpus

Each generator takes a correctly saved file and injects exactly one defect. The test asserts that the expected gate fails with the expected path.

| Corruption | Expected detection | Legacy case |
|---|---|---|
| Revert an edited cell to its O value (edit "lost") | G4 `intent_mismatch` | V-06 |
| Remove one child of a run's `rPr`; flip `b`; change a run color | G5 `undeclared` on `.rich` | V-02 (inverse) |
| Shift one row's content by one position after an insert | G5 on the transformed cell | V-04 |
| Remove `xml:space="preserve"` from text with a leading space | G5 `.rich` / `.value` | — |
| Add an undeclared prefix to `mc:Ignorable` | G3 | EX-04 |
| Strip the namespace from `x14ac:dyDescent` | G3 | EX-03 |
| Delete a chart part / media / printer settings | G1 / G6 | EX-01 |
| Delete the VBA relationship; change the content type to non-macro | G1 / G6 | EX-01 |
| Change the workbook password hash | G5 `wb.protection` | EX-05 |
| Drop `hidden` from `_xlnm.Print_Titles` | G5 `name[…]` | EX-07 |
| Move a merge range by one row | G5 `S!merge` | — |
| Off-by-one formula reference shift | G5 `rewrite_disagreement` | — |
| Change a shared string's text (one used by untouched cells) | G5 on every referencing cell | — |
| Change a numFmt / column width / sheet state / sheet order | G5 | — |
| Remove a CF rule; change a DV formula; change a table ref | G5 | — |
| Change one byte inside an unmodeled `extLst` | G5 canonical-XML fallback | — |
| Renumber relationship ids **without** updating the references | G1 (dangling) | — |

**Rule:** every verifier miss found later (in production, in evals, or in the parity harness) becomes a new corpus item before it is fixed.

**Differential fuzzing.** FsCheck generates random formulas (A1/absolute/mixed refs, ranges, whole rows/columns, 3D refs, names, functions) and random transforms. The engine's `Formula` shifter and the verifier's independent shifter must agree. Any disagreement is minimized and kept as a regression case.

### 2.5 Verifier specificity: equivalence corpus

- For each normalization rule (04 §5), generate a pair of variants and check them both ways (A→B, B→A): each must be classified `normalized` with the right rule id, and the save must not be blocked.
- **Exploratory input:** files saved by Excel itself (an untouched file re-saved through the oracle) and by LibreOffice. Their differences against the original are listed and triaged by hand into new rules or real differences. They are never auto-accepted.

### 2.6 Parity with the legacy server

- The same scenario scripts run against the Python server (`servers/excel`) and the C# server.
- The written files are compared semantically with `excel_verify`.
- Expected differences are listed in `tests/DocLoupe.Excel.Parity/expected-differences.md` with a reason each. Legacy rebuild artifacts are typical; C# behavior that differs on purpose (05) is the other main category.
- Legacy scenarios to port first: the no-edit round trip of fixtures 00–07, the supplemental 28-case workflow, row-height preservation, rich-text creation, and formula shifting.

### 2.7 Agent-level evals

This is the measure the project exists for. A suite of about 40 realistic tasks, run by an LLM agent against both servers. Each task has an **oracle checker**: a script that inspects the output file and decides whether it is correct.

**Metrics:**

| Metric | Definition | Target for the C# server |
|---|---|---|
| **Silent failure rate** | The agent reports success, but the checker says the file is wrong | **0** |
| Task success rate | The checker passes | ≥ legacy + 15 percentage points |
| Detected failure rate | The tool reported a problem and the agent surfaced it | Tracked |
| Tool calls per task | — | ≤ legacy |
| Tokens per task | Input + output | ≤ 70% of legacy |

**Example tasks** (with Vietnamese content where relevant):

- Recolor one word inside a cell that has 3–5 differently formatted runs; keep all other runs.
- Fix a typo in a 5-run cell without losing any formatting.
- Apply a fill to a range that contains rich cells, leaving the runs intact.
- Insert 3 rows into a table that has a totals row; check that the totals and the cross-sheet formulas still point to the right ranges.
- Rename a sheet that is referenced by formulas, defined names and a chart.
- Merge a header across columns that contain data (the agent must handle `MERGE_WOULD_DISCARD_DATA`).
- Set a VND number format and confirm that the display shows `1,250,000 ₫` under the en-US display locale (Q6), with `locale_dependent: true`.
- Edit one cell in a file whose sheet XML uses the `x:` prefix (the V-06 case).
- Edit one cell in a 500k-cell sheet.
- Add a list data validation and print titles to a `.xlsm`, and keep the macros.

### 2.8 Performance

The initial budgets are calibrated in spike S4. Benchmarks use BenchmarkDotNet, and large fixtures are generated with a fixed seed rather than committed.

| Scenario | Budget |
|---|---|
| `excel_open` on a 100 MB package (lazy) | < 3 s, < 300 MB RSS |
| `excel_read` page of 2,000 cells | < 200 ms |
| Single-cell `excel_apply` on a 1M-cell sheet (streaming) | < 1 s, memory independent of sheet size |
| `excel_save` with all gates, 1M-cell sheet touched | < 30 s |
| `excel_save` with 1 touched sheet out of 50 | Proportional to the touched sheet + package scan |

### 2.9 Fixtures

- **Location:** `tests/fixtures/` with a `manifest.json` giving the sha256, source, generator and features of each file.
- **Generator diversity matters**, because V-06 came from a namespace-prefix style the regex never saw. Fixtures must include files written by:
  - Excel (Windows and Mac);
  - LibreOffice;
  - Google Sheets export;
  - .NET Open XML SDK (`x:` prefixes);
  - openpyxl;
  - XlsxWriter;
  - typical ERP/report exporters.
- **Commit policy (Q3, decided: no).** The legacy external fixtures (`D:\data-test\excel-preservation-fixtures\sources\00–07`) are **never committed**. They run only in the local extended suite, located through `DOCLOUPE_EXCEL_FIXTURE_ROOT` and skipped when absent, as the legacy tests already do.
- **What CI uses instead:** committed equivalents that we create ourselves from synthetic content, covering the same features:
  - rich text with phonetics, table metadata, print area/titles: from the generator;
  - VBA and digital signatures: authored once in Excel on the developer machine from synthetic content.

  Only files we created may be committed.
- A deterministic C# fixture generator builds feature-specific workbooks (rich text with phonetics, shared/array formulas, tables with totals, CF/DV, print settings, drawings, signatures, VBA stub).

### 2.10 CI

| Job | Runner | Contents |
|---|---|---|
| Unit + op + corpus + matrix | GitHub-hosted, all 4 OSes | Everything except the Excel oracle |
| Parity | GitHub-hosted (Python available) | Legacy vs C# scenarios |
| LibreOffice oracle | GitHub-hosted Linux/Windows with LibreOffice installed | Render/recalc smoke tests |
| Excel oracle | **Developer machine with Office, run manually** (Q4, decided: no self-hosted runner; GitHub-hosted runners have no Office). Mandatory step of the release checklist; the report is attached to the release PR | `open_check`, render, recalc, display-string capture |
| Agent evals | Scheduled, with model API access | §2.7 metrics; a regression on silent failures fails the job |
| Release | Existing matrix, `dotnet publish` per RID (Linux inside the old-glibc container, 01 §12) | Binary smoke tests per platform |
| Clean-machine smoke | Containers/VMs **without** .NET, ICU or OpenSSL: Ubuntu 20.04, Debian 11, RHEL 8 (UBI minimal), current Ubuntu minimal; a fresh Windows 10 and Windows 11 VM without .NET or Office; macOS at the minimum supported version | Two modes: (a) through the npm launcher; (b) **direct binary with no Node installed**, using `--self-check` and then an MCP stdio session doing open/read/apply/save/verify on fixtures. Check that oracle absence is reported, not fatal. On Windows, **Windows Sandbox** (a pristine image) is a cheap local substitute for the clean VM. The harness already exists in `tests/clean-machine/windows-sandbox/` (scenario `legacy`); add a scenario for the C# tools. Baseline (2026-10-01): the legacy v1.1.2 binary passed all 13 steps in Windows Sandbox (no Python/Node/.NET/VC++ redist, networking disabled) |

## 3. Spikes

| ID | Question | Method | Exit criterion |
|---|---|---|---|
| S1 | Do Open XML SDK 3.x and the MCP C# SDK work under Native AOT? | Minimal server: open a workbook, read a sheet, return `structuredContent`. Collect trim/AOT warnings; measure binary size and startup on the 4 RIDs | AOT on all RIDs, or a documented decision to use the single-file fallback |
| S2 | Does DOM serialization preserve untouched content inside touched parts (prefixes, `mc:Ignorable`, `x14ac`, `extLst`, unknown elements)? Approach A (`SpreadsheetDocument` editable) vs B (PackageStore + detached part DOM) | Round-trip a touched sheet from every fixture; canonical-XML diff must be empty; check untouched parts byte-for-byte and `[Content_Types].xml` | Choose B unless A is proven byte-stable for untouched parts and content types |
| S2b | Can B write a touched part with the source's exact prefixes, root declarations and XML declaration, with part DOMs built **without** `System.IO.Packaging`? Does G3 reject EX-04? | Build part DOMs from raw XML; prefix-restoration pass (XmlReader → XmlWriter, URI → source prefix); re-run the S2 corpus plus Excel-authored files that carry `<?xml … standalone="yes"?>`; independent G3 check. Also check, using Excel, whether the `<ns0:Types>` fixtures open at all | Touched parts match O lexically except the edited elements on every fixture; EX-04 rejected; `ns0:` content types read without being rewritten |
| S3 | How do the MCP SDK and the target clients handle `structuredContent`, `outputSchema`, cancellation and image content? | Test tool against Claude Code, Claude Desktop and VS Code: which block reaches the model, whether cancellation arrives, whether images are shown | Envelope format fixed (03 §1.1) |
| S4 | Streaming editor performance and the DOM threshold | Generated 100k / 1M / 5M-cell sheets; single-cell and range edits; memory and time | Budgets in §2.8 confirmed or revised; `DOCLOUPE_EXCEL_DOM_MAX_BYTES` default chosen |
| S5 | Excel COM oracle: render (`CopyPicture` vs `ExportAsFixedFormat`), recalc, **detecting a repair prompt**, process isolation, macro security | Helper process prototype; deliberately broken files | Reliable repair detection, or `open_check` dropped |
| S6 | LibreOffice oracle across OSes: render a range (print area on a temp copy) directly to PNG; recalc | Prototype on the 3 OSes | Fidelity notes written; timeouts calibrated; no rasterizer needed in the server |
| S7 | Display strings and the rich-run font merge rule | Compare `ExcelNumberFormat` with Excel on ≥ 300 pairs; render cells whose run `rPr` omits properties | Display mismatches documented or fixed; merge rule in 02 §3.4 confirmed |
| S8 | Is the `excel_apply` schema too large? | Measure the serialized `inputSchema`, compare with the legacy 112 schemas; mini agent eval of single vs split apply | Decision on 03 §8 |
| S9 | Does the binary run on a clean machine with no runtime or extra libraries (01 §12)? | Build with `InvariantGlobalization=true` and managed hashing. Inspect dynamic dependencies (`ldd`, `otool -L`, `dumpbin /dependents`). Run the smoke suite on the clean-machine matrix (§2.10). Check grapheme segmentation, NFC normalization and case-insensitive `Đ`/`đ` matching in invariant mode | Dynamic dependencies limited to the OS C library family (Linux), system frameworks (macOS) and Windows system DLLs; smoke suite green on every clean image; glibc floor documented |

**S1/S2 Windows spike result (2026-10-01):** `../spikes/S1S2/REPORT.md`. Native AOT `win-x64` produced a working MCP read tool with `structuredContent`, 41,724,928-byte executable and 87.8 ms median start-to-initialize (10 local runs); `dumpbin` dependencies were recorded. Linux/macOS RIDs and clean-machine validation remain open. S2 selects B because A (`System.IO.Packaging`) cannot open 3/8 fixtures and gives no control over content types; the `xl/workbook.xml` rewrite first attributed to A was an `AutoSave` artifact (with `AutoSave = false` A kept all untouched entries byte-identical). The 8/8 byte-identical untouched entries under B do **not** mean the touched sheet passed: Open XML SDK rewrote the default namespace to `x:`. Under the decided G3 lexical-fidelity rule (04 §3.3) this must be fixed by S2b before P2.

## 4. Roadmap

| Phase | Scope | Exit criteria |
|---|---|---|
| **P0** Spikes | S1–S9, S2b | All spike exits met; this design updated with the results |
| **P1** Read + verify | `excel_open` (read-only), `excel_peek`, `excel_read`, `excel_find`, `excel_inspect`, `convert_to_markdown`, `excel_verify` (G1–G3, G5, G6; plus G4/G7 when the caller supplies `declared.effects` or `assert`) | Corruption corpus 100% for the modeled facets; equivalence corpus passes; read parity with legacy. **Useful early, within limits:** on files saved by the Python server it finds undeclared losses (EX-01/EX-03/EX-05 class). A *lost* edit (V-06) leaves W equal to O, so a plain two-file comparison cannot see it; it is found only when the caller passes the expected values via `assert` or `declared.effects` (03 §6.1). V-04-class losses after structural edits need `declared.transforms` |
| **P2a** Thin write slice | `excel_open`, `excel_read` (cells view), `excel_apply` with **only `set_value`**, `excel_save` (staging → G1–G5 → readback from the written file), `excel_close`. No rich text, no structural edits, no other ops | S2b passed; `set_value` mutation matrix green on all fixtures, including prefixed-namespace sheets (the V-06 shape) saved correctly; corruption corpus for G1–G5 at 100%, including a forced lost edit blocked by G4 |
| **P2b** Cell editing | Remaining cell-content, rich-text and style ops; `excel_undo`; G6–G7; `excel_status`, `excel_create` | Op tests and mutation matrix green for these ops; agent-eval subset at 0 silent failures |
| **P3** Structure | Rows, columns, merges, sheet ops; transforms; `Formula` module + independent shifter | Differential fuzzing clean after the agreed run budget; matrix green |
| **P4** Features | Tables, CF, DV, names, hyperlinks, comments, printing, views, protection, workbook/document properties | Matrix green; EX-05/EX-07/EX-08/EX-10 scenarios pass |
| **P5** Drawings + package | Drawing ops, expert package ops, `excel_export` | Matrix green; EX-01 scenario passes |
| **P6** Oracles + hardening | `excel_render`, recalc, `open_check`; performance budgets; full agent eval; release as **`excelnext`** | §2.7 targets met; budgets met; smoke tests on 4 platforms |
| **P7** Switch | `excel` points to the C# binary. The Python server stays available for one release under a launcher-compatible name (e.g. `excellegacy`), then is removed | Every row of the 05 mapping is implemented or explicitly dropped; no open critical defects; **S1 exit met on all four RIDs and the S9 clean-machine suite green**, measured on clean images rather than inferred from the Windows-host spike |

## 5. Risks

| Risk | Impact | Mitigation |
|---|---|---|
| Open XML SDK or the MCP SDK not AOT-ready | Larger binary, slower start | Single-file fallback (01 §9); S1 early |
| The binary needs something missing on the user's machine (ICU, OpenSSL, a newer glibc, a native library) | Server does not start, or fails mid-operation | Invariant globalization, managed hashing, old-glibc build container, no side-by-side native libraries, clean-machine CI (01 §12, S9) |
| SDK serialization changes untouched content inside a touched part | False positives, or real changes | Canonical comparison + registered rules only; S2 decides the approach |
| `excel_apply` schema too large for agents | Lower selection and parameter accuracy | Split into three tools with the same semantics (03 §8); decided by eval |
| Clients ignore `structuredContent` | Agent sees only text | Compact JSON text block always present |
| Formula-shifting edge cases (structured refs, 3D, spill, external, `INDIRECT`) | Wrong references | Independent shifter, fuzzing, required gaps instead of silent passes, `TEXT_REFERENCE_NOT_SHIFTED` warning |
| No Excel in CI (Q4: no self-hosted runner) | Excel-specific regressions (open/repair, display) are found only at release time | LibreOffice oracle in CI; manual Excel-oracle run on the developer machine as a release gate; committed Excel-captured display table catches display regressions in CI |
| Display depends on locale | Agent sees a different string than a user with other regional settings | Default `en-US` (Q6); `display_locale` in every read; `locale_dependent` flag; configurable |
| Run-font merge rule wrong | Wrong effective style in readback | S7 before P2 |
| Three-way verification on huge files is slow | Save latency | Streaming comparison, per-sheet parallelism, untouched parts by hash |
| Scope creep into pivots, slicers, etc. | Delay | Preserve-and-verify only (01 §1) |
| Agents and prompts that use legacy tool names | Broken workflows at the switch | `convert_to_markdown` kept; mapping table (05); clear server instructions; one-release overlap |

## 6. Open questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | Should `excel_save` default to `mode: "overwrite"`, or require an explicit mode? | **Decided 2026-10-01 (design review): explicit `mode` required**, no default (03 §5) |
| Q2 | Should every `accept` item require a `reason` that is recorded in the report? | **Decided 2026-10-01: yes**, for `accept` and `override`. `accept` is limited to gaps, `override` to G5 undeclared differences (04 §4) |
| Q3 | Can the external fixtures in `D:\data-test\excel-preservation-fixtures` be committed (privacy/licensing)? | **Decided 2026-10-01: no.** Local extended suite only; CI uses fixtures we create (§2.9) |
| Q4 | Is a self-hosted Windows runner with Office available? | **Decided 2026-10-01: no.** Excel oracle runs manually on the developer machine as a release gate (§2.10) |
| Q5 | One `excel_apply`, or three apply tools? | Decide with S8 + evals |
| Q6 | Default display locale: `en-US` or `vi-VN`? This matters for Vietnamese users: thousands separators and date order differ | **Decided 2026-10-01: `en-US`.** Configurable; `display_locale` reported on every read (02 §2.2) |
| Q7 | Keep legacy tool names as thin aliases during the transition? | No, except `convert_to_markdown`; it keeps the surface lean and avoids two semantics for the same name |
| Q8 | Should a session ledger survive a server upgrade? | Only within the same ledger schema version; otherwise refuse `resume` with a clear error |
