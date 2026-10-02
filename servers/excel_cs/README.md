# excel_cs — C# port of the Excel MCP server (design)

**Status:** design, spikes, a source-preserving P2a writer, and partial P2b cell editing; not yet production-ready. **Scope:** the Excel server only (`servers/excel`). This port replaces the current Python implementation (112 tools, `server_version` 1.1.2) incrementally.

These documents are for the engineers who will implement the C# server. Every decision is mapped back to the legacy server's behavior, so each one answers three questions: **how the old server does it, where it goes wrong, and what the new one must do.**

## Why port

The primary goal of the tool is that an agent **knows whether its edit is right or wrong**, and that **the saved file matches what the agent believes it wrote**. The legacy server cannot deliver this because of its foundations:

- **It writes through openpyxl, which loads the workbook into its own model and rewrites everything.** Whatever openpyxl does not model is dropped. The code compensates with 17 `_inject_*` passes, `_restore_missing_package_parts`, and a regex-based "patch only the edited cells" fast path.
- **It edits XML with regex and stdlib `xml.etree`.** About 68 regex calls in `core.py` match XML tags. Together these lose or rename namespace prefixes and break `mc:Ignorable`.
- **Its verifier is too weak.** It compares only the original file with the saved file. It exempts the edited region from checks, and it **never checks that the requested edit actually reached the file**. Verification is also off by default.

Confirmed defects (details in [docs/05-legacy-mapping.md](docs/05-legacy-mapping.md#known-defect-register)):

- **V-06:** when the sheet XML uses a namespace prefix (`<x:c>`), the edit is silently lost, yet both save and verification report success.
- **V-04:** inserting or deleting a row makes the verifier treat the whole sheet as "requested", so it can detect nothing on that sheet.
- **V-02 / V-07:** false positives after a full rebuild (`<b val="1"/>` becomes `<b/>`, `14` becomes `14.0`).
- **EX-01 … EX-10:** see the 2026-09-30 audit report.

## Key decisions

| # | Decision | Replaces (legacy) |
|---|---|---|
| D1 | .NET 10 (LTS), C#; the official MCP C# SDK over stdio | Python + FastMCP |
| D2 | **System.Xml part DOM + source-byte-preserving splice** is the P2a writer core. Untouched part contents and bytes outside declared edits stay identical. Open XML SDK is a separate detached G2 schema adapter, not the writer or package loader (01 §5.1; P2a SDK-role spike) | openpyxl rebuild, regex patching, `xml.etree` |
| D3 | The verifier uses an **independent reader** (`ZipArchive` + `XmlReader`) that shares no code with the writer | Verifier built on ElementTree, sharing assumptions with the writer |
| D4 | **Three-way check**: original (O), intent (I), written file (W) | Two-way, O vs W |
| D5 | Verification is **always on**. A failing gate means nothing is written. The single exception: G5 differences the user explicitly overrides, which produce `saved_with_overrides`, never "verified" | `verify_preservation=False` by default |
| D6 | All mutations go through **one** tool, `excel_apply`. It is transactional (one failing op rolls back the whole batch) and supports `dry_run` and `expect` | About 70 separate mutation tools, not atomic across tools |
| D7 | **A1 addressing only**, with 1-based Excel row numbers | A mix of 0-based `r1/c1`, A1, and exclusive `end_row` |
| D8 | Rich text: a readable and writable markup, edits **by content `match`**, offsets counted in graphemes | Code-point offsets that split Vietnamese diacritics |
| D9 | Results use the **same representation as the input**: effective formatting, display strings, markup | Internal views, JSON embedded in text |
| D10 | Three states: `verified` / `failed` / `unverified`. Never report "passed" for something that was not checked | `preservation_ok: true` while a whole sheet was skipped |
| D11 | Optional external oracles: rendering, recalculation, "does Excel ask to repair this file" | Only `excel_capture` via LibreOffice |
| D12 | One self-contained single-file binary per platform with the **same asset name** `excel-tools`. Users install no runtime or libraries, and the npm launcher does not change (01 §12) | PyInstaller (self-contained too, but extracts to a temp directory on every start) |

## Documents

| File | Contents |
|---|---|
| [docs/01-architecture.md](docs/01-architecture.md) | Stack, project layout, session model, apply and save pipelines, concurrency and cancellation, limits, packaging and release |
| [docs/02-data-model.md](docs/02-data-model.md) | Addressing, value model, display values, effective style, rich text markup grammar, graphemes and NFC, semantic paths, hashes |
| [docs/03-tools.md](docs/03-tools.md) | The 15 new tools: detailed input/output, response envelope, error codes, and the `excel_apply` op catalog |
| [docs/04-correctness.md](docs/04-correctness.md) | The correctness contract with the agent: invariants, save gates, normalization rules, structural edits, oracles, worked scenarios |
| [docs/05-legacy-mapping.md](docs/05-legacy-mapping.md) | All 112 legacy tools mapped to new tools/ops; conventions, error codes, environment variables; known defect register |
| [docs/06-testing-and-roadmap.md](docs/06-testing-and-roadmap.md) | Test strategy (including verifier sensitivity testing and agent-level evals), spikes, phased roadmap, risks, open questions |

**Suggested reading order:** README → 04 (why) → 03 (what) → 02 (representation) → 01 (how) → 05 (legacy mapping) → 06 (order of work).

## Glossary

| Term | Meaning |
|---|---|
| **O / I / W** | The original file at open / the intended state after the ops / the file actually written |
| **Op** | One mutation inside `excel_apply`, e.g. `set_value`, `rich_style`, `insert_rows` |
| **Effect** | A change an op declares it will cause: semantic path, value before, predicted value after |
| **Transform** | A coordinate mapping caused by a structural op (row/column insert or delete, sheet rename). Applied to O before comparing it with W |
| **Ledger** | The per-session log of applied ops with their effects and transforms, numbered by revision |
| **Gate** | One save-time check, with status `verified` / `failed` / `unverified` |
| **Oracle** | An external checker: Excel via COM, or headless LibreOffice |
| **Facet** | One aspect of a cell: `value`, `formula`, `rich`, `style.font.bold`… |
| **Display value** | The string Excel shows in the cell after applying the number format |
| **Effective style** | The formatting a user actually sees, after merging the xf, the cell font, the run `rPr`, and resolving theme colors to RGB |

## P2a working slice

The production `DocLoupe.Excel.slnx` now contains `Model`, `Package`, `Engine`, independent `Verify`, SDK-only `Schema`, `Server`, and their test projects. `dotnet test servers/excel_cs/DocLoupe.Excel.slnx` runs on the six repository-generated fixture shapes. Local, non-distributable sources can additionally be tested by setting `DOCLOUPE_P2A_LOCAL_FIXTURES` to `D:\data-test\excel-preservation-fixtures\sources` before `dotnet test`; CI never reads that path. Run `dotnet run --project servers/excel_cs/tools/SchemaOrder/SchemaOrder.csproj` to regenerate/check element-order evidence against the pinned Open XML SDK.

The MCP server offers `excel_open`, `excel_status`, `excel_undo`, `excel_read` (targeted cells), `excel_apply` (`set_value` on a cell or rectangular range, rectangular `set_values`, normal `set_formula`, bounded `fill`, and value-only `clear`), `excel_save` (`mode: "copy"` only, new path), and `excel_close`. All edits are staged; G1–G6, including G4 readback and the independent P2a G6 protection of advanced parts, block output on failure. String cells use shared strings unless an existing inline-string cell retains inline storage. `fullCalcOnLoad` is enabled on value changes; `calcChain` is removed only when an existing ordinary formula is overwritten. Signed packages and formula-group edits need their later-phase policies and must be rejected rather than silently rewritten. P2a rejects unsupported encodings and worksheets above 32 MiB instead of falling back to full serialization.

P2b now includes `excel_undo(session, base_revision, to_revision)` by replaying only retained cell-content revisions, and `excel_status(session?)` with session listing, the last 20 revision summaries, source-change detection and busy state during save. Copy saves leave the session dirty until undo or close; `saved_revision` stays at 0 because copy-save does not update the opened source. Oracle fields report `unavailable` until oracle integration is implemented, regardless of installed programs; build metadata reports `development`/`unknown` unless embedded in the assembly. A further P2b slice adds fail-closed `excel_save.assert` checks for sheet-qualified cell `equals.value` (blank, string, boolean, supported error tokens, bounded numeric lexemes) and `equals.formula`, evaluated on the staging file as G7. Unsupported assertion facets (`display`, `rich`, `style`, `except`) reject the request rather than being silently ignored; standalone `unchanged: true` compares the selected cell's XML subtree, resolved value and referenced shared-string `<si>` against the opened source, including an absent cell remaining absent. It does not claim that unrelated cells or package metadata remain unchanged; those are checked by the other save gates. `equals.value` on formula cells verifies the typed cache, not the formula's mathematical result. `set_formula` supports ordinary formulas with `cache: "clear"` or `cache: "keep"` on an existing formula, rejects shared/array groups and other cache modes, and independently verifies retained caches against the source after save. `set_value` broadcasts one value to a rectangular range of at most 500 cells; `{"error":"#N/A"}` and the error tokens recognized by the legacy formula-cache writer store as `t="e"`. `fill` reuses this bounded path for constant values and row-major exact decimal `series: {start, step}` (15 significant digits per result, bounded magnitude and exponent); `pattern_from` remains unsupported. `set_values` accepts a rectangular `values` matrix with a matching range or top-left anchor, expands to at most 500 cell edits, and applies in one revision; other bulk options remain unsupported. `clear` accepts only `what: ["values"]` (or omitted): `remove_cells: false` keeps empty styled cells; `remove_cells: true` removes cells but retains rows. It does not create absent cells, and removes `calcChain` only when clearing an existing formula. Other `clear` facets fail closed. No redo or other P2b ops are available yet. The P2a audit aligned sheet-qualified addresses (they override the default sheet), explicit `as_text` shared-string storage, phonetic-free value readback, merged-cell origin safety, and G4's missing-target report with 02/03 and the legacy reader/edit guard. Broader 03 schemas, persistent ledgers, full G6 policy for deliberate advanced-part changes and signature invalidation, the remaining G7 facets, and model-visible S3 client forwarding remain outside this slice.

**Scope of evidence:** local Windows tests and the official C# SDK stdio spike passed for this increment; the new commit has not yet been exercised by remote cross-platform CI. The S3 report does not establish what Claude Code, Claude Desktop or VS Code forwards to a model. The current verifier's byte-span checks accept writer-declared spans, and the P2a G5 cell-facet check covers edited cells; broader semantic verification and the remaining gates/oracles are not implemented. Do not treat this as production-ready or as the complete 01 §3 roadmap.
