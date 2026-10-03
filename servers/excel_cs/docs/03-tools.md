# 03 — Tool surface

112 legacy tools collapse into **15 tools**. All mutations go through `excel_apply` as a list of typed ops. The full per-tool mapping is in [05-legacy-mapping.md](05-legacy-mapping.md).

| # | Tool | Session | Purpose |
|---|---|---|---|
| 1 | `excel_open` | creates | Open a workbook for reading/editing; returns an overview |
| 2 | `excel_create` | creates | Create a new workbook (optionally from a template) |
| 3 | `excel_status` | reads | Session state, ledger, oracle availability; lists sessions when called without one |
| 4 | `excel_close` | ends | Close a session; refuses to drop unsaved work unless told to |
| 5 | `excel_read` | reads | Read cells/ranges as cells, values, Markdown or full detail |
| 6 | `excel_find` | reads | Find cells by text/regex/value/formula/style |
| 7 | `excel_inspect` | reads | Read workbook/sheet objects: tables, names, CF, DV, drawings, print, views, package… |
| 8 | `excel_apply` | mutates | Apply a transactional batch of ops; supports `dry_run` and `expect` |
| 9 | `excel_undo` | mutates | Roll the session back to an earlier revision |
| 10 | `excel_save` | writes file | Save with mandatory verification gates |
| 11 | `excel_verify` | none | Validate a file, or compare two files (for files produced elsewhere) |
| 12 | `excel_render` | optional | Render ranges to PNG via an oracle so a multimodal agent can see the result |
| 13 | `excel_export` | reads | Export images, package parts, or a range as CSV to disk |
| 14 | `excel_peek` | none | Quick read-only look at a file without opening a session |
| 15 | `convert_to_markdown` | none | One-shot conversion to Markdown (legacy-compatible signature) |

Types below use a TypeScript-like notation. `?` marks an optional field; `= x` gives the default. The C# implementation uses `record` DTOs with `[JsonPolymorphic]` on `"op"`, and JSON source generation.

## 1. Common conventions

### 1.1 Response envelope

Every tool returns `structuredContent` that conforms to its `outputSchema`, plus the same JSON serialized compactly in a `TextContent` block. The MCP spec recommends the text block for backward compatibility; spike S3 checks which block each target client actually forwards to the model.

```ts
type Ok<T> = {
  ok: true
  session?: string            // when the tool works on a session
  revision?: number           // revision after the call
  data: T
  warnings: Warning[]         // never omitted; empty array when none
  next_cursor?: string        // paginated results
  metrics?: Metrics           // when include_metrics = true
}
type Warning = { code: string, message: string, path?: string }

type Err = {
  ok: false
  error: {
    code: string              // stable, see §5
    message: string           // human-readable, one sentence
    details?: object          // machine-readable specifics (op index, expected/actual, candidates…)
    hint?: string             // the next action that would likely succeed
    retryable: boolean
  }
}
```

- **Keys:** all keys are `snake_case`, and no two keys in one object may differ only by case. The legacy server emits both `vAlign` and `valign` in one cell view, which case-insensitive JSON consumers (Windows PowerShell 5.1 `ConvertFrom-Json`, .NET with `PropertyNameCaseInsensitive`) reject or silently merge (V-11). A unit test walks every `outputSchema` to enforce this rule.
- Errors are returned with `isError: true`.
- A **blocked save** is an error (`SAVE_BLOCKED`), and the full gate report goes in `details`. The agent must never be able to read a blocked save as a success.

### 1.2 Shared input fields

| Field | Meaning |
|---|---|
| `session` | Session id returned by `excel_open` / `excel_create` |
| `sheet` | Default sheet for addresses that are not sheet-qualified |
| `base_revision` | Required by `excel_apply`. Must equal the current revision, otherwise `REVISION_CONFLICT` |
| `cursor` / `max_cells` / `max_results` | Pagination |
| `include_metrics` | Adds the `metrics` block |

### 1.3 Server instructions (MCP `instructions`)

```
DocLoupe Excel — verified editing of .xlsx/.xlsm/.xltx/.xltm.
Workflow: excel_open → excel_read/excel_find/excel_inspect → excel_apply (with expect, optionally dry_run)
→ excel_save → read the save report. Addresses are A1 with 1-based rows (e.g. Sheet1!B5, 5:7, C:D).
Every op in one excel_apply call refers to the state you read (base_revision); the server handles
shifting. Rich text: use markup (<r b color="FF0000">text</r>) and edit by "match", not by offsets.
excel_save requires an explicit mode. A save is fully successful only when status is "verified";
"verified_with_gaps" lists what could not be checked; "saved_with_overrides" means known differences
were written on purpose: use override only with the user's explicit consent.
SAVE_BLOCKED means nothing was written; read error.details.gates. Never assume an edit landed without
checking the readback in the apply/save result.
```

## 2. Session and file tools

### 2.1 `excel_open`

```ts
input  { path: string, resume?: string, read_only?: boolean = false, include_metrics?: boolean }
output data {
  path: string, format: "xlsx"|"xlsm"|"xltx"|"xltm",
  fingerprint: { sha256: string, size: number },
  sheets: { name: string, index: number, kind: "worksheet"|"chartsheet"|"dialogsheet",
            state: "visible"|"hidden"|"veryHidden", used_range: string|null,
            tables: string[], merged_ranges: number }[],
  workbook: { date1904: boolean, calc: { mode: string, full_calc_on_load: boolean },
              defined_names: number, active_sheet: string },
  features: { vba: boolean, signatures: boolean, pivots: number, slicers: number, timelines: number,
              external_links: number, charts: number, images: number, shapes: number,
              comments: number, threaded_comments: number, custom_xml: boolean, data_model: boolean },
  resumed_revision?: number
}
```

- **Warnings:**
  - `SIGNED_WORKBOOK`: any edit will invalidate the signatures, and saving will require `allow_signature_invalidation`.
  - `VBA_PRESENT`, `EXTERNAL_LINKS`, `NFD_TEXT_PRESENT`.
  - `BASELINE_SCHEMA_ERRORS`: the source already has schema errors. They are recorded so they are not blamed on later edits.
- **Errors:** `FILE_NOT_FOUND`, `UNSUPPORTED_FORMAT`, `PACKAGE_INVALID`, `LIMIT_EXCEEDED`, `PATH_NOT_ALLOWED`, `RESUME_FINGERPRINT_MISMATCH`.

### 2.2 `excel_create`

```ts
input  { format?: "xlsx"|"xlsm"|"xltx"|"xltm" = "xlsx", target_path?: string,
         sheets?: string[] = ["Sheet1"], active_sheet?: string,
         template_path?: string,            // start from an existing workbook/template (content copied)
         vba_from?: string,                 // .xlsm/.xltm whose vbaProject.bin is copied (macro formats only)
         document_properties?: DocProps }
output data { same shape as excel_open, plus new: true, default_path: string }
```

A new workbook has no user-supplied O before creation. The current slice publishes the base workbook first and uses it as O for subsequent edits; G1–G7, including G5, run on the first edited save.

**Current P2b create slice:** a distinct, nonexistent `target_path` with `.xlsx`, `.xlsm`, `.xltx` or `.xltm` extension is required; optional `format` must match that extension. Without a template, it creates a minimal workbook with `sheets` (default `["Sheet1"]`) and optional `active_sheet`, checks OPC relationships and detached schema validation, then publishes atomically as a clean session. Macro-capable blank formats contain no VBA. With `template_path`, both files must use the same OOXML extension; it copies the template byte-for-byte via staging, verifies the source fingerprint, and opens the copy as a clean session without sheet changes. Copying existing macro parts is byte-preserving, but editing signed packages still fails G6. New workbooks optionally accept `document_properties.core` with string `title`, `subject`, `creator`, `description`, `keywords`, `category`, `contentStatus`, or `lastModifiedBy` (up to 4096 characters each). The server writes an OPC core-properties part and relationship, validates this limited subset independently (unsupported core fields remain G2 gaps), and preserves it byte-for-byte on subsequent cell edits. Template creation rejects document-property changes. `document_properties.app`/`custom`, dates and other core fields, VBA copied from a separate workbook, and an implicit target path remain unsupported; this is not yet the complete creation contract.

### 2.3 `excel_status`

```ts
input  { session?: string }
output data (with session) {
  path: string, revision: number, saved_revision: number, dirty: boolean, read_only: boolean,
  source_changed_on_disk: boolean,
  busy: { operation: string, since: string } | null,
  ledger: { revision: number, op_count: number, summary: string }[],   // last 20
  server: { version: string, commit: string,
            oracles: { excel: "available"|"unavailable", libreoffice: "available"|"unavailable" } }
}
output data (without session) { sessions: { session: string, path: string, revision: number, dirty: boolean }[], server: {...} }
```

### 2.4 `excel_close`

```ts
input  { session: string, discard_unsaved?: boolean = false }
output data { closed: true, discarded_revisions: number }
```

Fails with `UNSAVED_CHANGES` when `revision > saved_revision` and `discard_unsaved` is false. The legacy server closes silently.

### 2.5 `excel_undo`

```ts
input  { session: string, base_revision: number, to_revision: number }
output data { revision: number, discarded: number[] }
```

- Replays the ledger from O up to `to_revision`. There is no redo.
- Undoing below `saved_revision` is allowed; the session then becomes dirty relative to the file on disk.

## 3. Read tools

### 3.1 `excel_read`

```ts
input {
  session: string, sheet?: string,
  target?: string | string[] | Target,          // default: the sheet's used range
  view?: "cells" | "values" | "markdown" | "full" = "cells",
  include?: ("rich"|"formula"|"style"|"comment"|"hyperlink"|"validation"|"cf")[]
            = ["rich", "formula", "style"],      // "cells" view only
  style_detail?: "sparse" | "full" = "sparse",
  skip_empty?: boolean = true,
  max_cells?: number = 2000, cursor?: string
}
output data {
  sheet: string, range: string, view: string, range_hash: string,
  display_locale: string,         // locale used for every `display` string (default "en-US")
  cells?: Cell[],                 // view = cells | full (see 02 §8)
  rows?: Value[][],               // view = values (typed values; dates as {"date": …})
  markdown?: string,              // view = markdown (A1 headers, Excel row numbers, † rich, ƒ formula)
  merged?: { range: string, origin: string }[]
}
```

| View | Use it for | Contains |
|---|---|---|
| `markdown` | Getting oriented, reading bulk content | Plain table, cheapest in tokens |
| `values` | Exact values in a block | 2D typed values |
| `cells` | Before editing | Values, display strings, rich markup, formulas, sparse style, hash |
| `full` | Debugging formatting | Everything, plus `xf_id`, `text_form`, color sources, formula attributes |

**Current P2b read slice:** `excel_read` accepts a sheet-qualified or default-sheet A1 cell, bounded rectangular target (`A1:C3`), or nonempty array of up to 500 such strings. Omitting `target` reads the first sheet's (or selected sheet's) independently inferred explicit-cell used range, including an empty sheet, subject to the same 500-cell bound. It returns only existing cells by default in `cells` view from the current revision (`skip_empty=false` includes addressable `blank` placeholders in target order), and rejects reversed, malformed, cross-sheet, or >500-cell ranges. The `values` view accepts one explicit cell or rectangular A1 range (at most 500 cells), including `null` for empty coordinates, typed numbers, text, booleans and errors; formulas return typed cached results or `null` for a missing cache. The `markdown` view accepts the same bounded rectangle and reports absolute column headers/row numbers, cached values with `ƒ =formula` annotations, and `†` for detected rich runs; it escapes table separators and HTML-sensitive content. Neither view applies `skip_empty` because missing positions remain visible. Date-format resolution, full used-range semantics (including row-only formatting), multiple targets in these views, hashes, pagination, rich/style projections, full read metadata and the `full` view remain future work.

### 3.2 `excel_find`

```ts
input {
  session: string,
  query: { text?: string, regex?: string, value?: Value, formula_contains?: string,
           style?: StylePredicate },                 // e.g. { "font": { "bold": true } }
  in?: "display" | "value" | "formula" | "rich" | "comment" = "display",
  scope?: { sheet?: string, target?: string },       // default: whole workbook
  case_sensitive?: boolean = false, normalize?: "nfc" | "none" = "nfc",
  max_results?: number = 100, cursor?: string
}
output data {
  matches: { addr: string,            // sheet-qualified, e.g. "Sheet1!B5"
             display: string, value: Value, hash: string,
             span?: [number, number],  // grapheme range of the hit inside the searched text
             snippet?: string }[],
  total_scanned: number
}
```

**Current P1 find slice:** `excel_find` accepts an optional `scope`: a single A1 cell/range on one sheet (`scope.sheet` or a qualified target), a whole sheet, or a whole workbook when omitted. Unscoped searches use the independent reader's explicit-cell used ranges, and fail closed if any range or the total exceeds 500 cells. Exactly one of `query.text`, `query.regex`, `query.formula_contains`, or `query.value` is required. `formula_contains` searches raw formula text as a literal substring, regardless of the default `in: "value"`; other `in` values beyond `value`/`formula` fail closed. Text/regex search is limited to raw `value` or `formula` text, with optional NFC normalization, case sensitivity, a 100 ms regex timeout and at most 100 results. `query.value` requires `in: "value"` (the default) and compares typed numbers, strings, booleans, errors, or blank values; formula cells use the typed cache without evaluating. String values honor `case_sensitive` and `normalize`; `null` also matches missing cells within the bounded search range. Output marks `partial: true`, reports addresses, typed values (including formula caches without recalculation), raw formulas and the count scanned; display, hash, grapheme spans, cursor, large-workbook paging and all other query facets are not available yet and fail closed. In this slice the `in` default is `value`, not the final contract's `display`.

### 3.3 `excel_inspect`

```ts
input  { session: string, aspect: Aspect, sheet?: string, id?: string, options?: object }
Aspect = "workbook" | "sheet" | "styles" | "tables" | "names" | "conditional_formats" | "validations"
       | "drawings" | "comments" | "hyperlinks" | "package" | "part"
```

| Aspect | Returns | Replaces (legacy) |
|---|---|---|
| `workbook` | workbookPr, calcPr, protection (flags and algorithm, never secrets), workbook views, document properties | `excel_get_workbook_semantics`, `excel_get_workbook_views` |
| `sheet` | state, properties, sheet views and freeze, print (page setup, options, header/footer, breaks, print area, titles), protection, protected ranges, ignored errors, autofilter, column/row property summary, merges | `excel_get_sheet_semantics`, `excel_get_sheet_views` |
| `styles` | Named styles, dxfs, table styles (effective) | — |
| `tables` | Tables with ref, columns, totals, style flags, filters | `excel_list_tables` |
| `names` | Defined names with scope, value, hidden, comment, built-in flags | `excel_list_defined_names` |
| `conditional_formats` | Rules with stable `cf_` ids, sqref, type, priority, effective dxf style, formulas | `excel_get_conditional_formats` |
| `validations` | Rules with stable `dv_` ids | — |
| `drawings` | Images/charts/shapes with `dr_` ids, A1 anchors with offsets, names, alt text, shape text as markup, chart series references | `excel_get_shapes` |
| `comments` | Legacy and threaded comments (read-only for threaded) | — |
| `hyperlinks` | All hyperlinks on a sheet | — |
| `package` | Parts with content types, sizes, hashes, relationship edges (bounded) | `excel_list_package_parts` |
| `part` | One part's content: `mode: "xml"\|"text"\|"base64"`, `offset`, `max_bytes` | `excel_read_package_part` |

## 4. `excel_apply`

### 4.1 Request and response

```ts
input {
  session: string,
  base_revision: number,
  sheet?: string,
  ops: Op[],                                   // 1..500 ops
  dry_run?: boolean = false,
  address_mode?: "base" | "sequential" = "base",
  expert?: boolean = false,                    // required for package ops and style_xf
  return?: "diff" | "diff+readback" = "diff+readback",
  max_diff_items?: number = 200
}

Op = { op: string, label?: string, sheet?: string, target?: string | Target, expect?: Expect, ...fields }

Expect = {                 // all optional; every given field must match the base state
  hash?: string            // strongest: exact content the agent read
  value?: Value, display?: string, text?: string, formula?: string, rich?: string,
  empty?: boolean, style?: StylePredicate
}

output data {
  dry_run: boolean,
  revision_before: number, revision_after: number,           // equal when dry_run
  results: { index: number, label?: string, op: string, status: "applied" | "planned",
             resolved: string,                               // e.g. "Sheet1!B12"
             moved_from?: string }[],                        // e.g. "Sheet1!B10" in base coordinates
  diff: { id: string, path: string, before: any, after: any }[],      // capped by max_diff_items
  diff_summary: { facets_changed: number, cells_touched: number,
                  by_category: Record<string, number>, truncated: boolean },
  transforms: { sheet: string, kind: string, at: string, count: number }[],
  readback: Record<string, Cell>,     // touched cells in the agent's representation, capped
  calc_effects: { full_calc_on_load: boolean, stale_formulas: number }
}
```

**Transactionality:** if any op fails, nothing is applied. The error names the op (`details.index`, `details.label`, `details.pointer` as a JSON pointer into the request), and for `PRECONDITION_FAILED` it includes `expected` and `actual`.

**Current P2b precondition slice:** `expect.value`, `expect.formula`, `expect.empty`, or their combinations on a single-cell `set_value`/`set_formula`/`clear`, or uniformly across every cell of a bounded rectangular `set_value`/`set_values`/`fill`/`clear`, are supported. They are checked against the session state at `base_revision` before any operation in the batch, with the independent cell reader and the same typed-value/formula comparison as `excel_save.assert.equals` (including formula caches, not formula evaluation). `expect.empty` considers a missing cell or a styled blank cell empty, but never a formula, even if its cache is empty; `false` requires content or a formula. `expect.formula` requires an existing formula and accepts a leading `=`. A mismatch returns `PRECONDITION_FAILED` with `details.index` (**original op index**, even if an earlier range expanded into multiple cells), `target`, `expected` and `actual`, leaving revision and cells unchanged. Other `expect` facets and per-cell expectation matrices fail closed; hashes, `label` and full diff remain future work. The current `dry_run` slice reuses the same preconditions and candidate writer, returning `intent` and `changed_parts` relative to the source (including earlier pending edits), with equal `revision_before`/`revision_after`; it does not change the session. Both dry-run and apply also return a bounded `readback` keyed by `Sheet!A1` for distinct cells touched by the current batch (at most 500): each entry uses the existing cell reader representation, or `null` if the cell was removed. This is a candidate readback, not a save-gate verification result; the full diff/results/transforms envelope above remains future work.

### 4.2 Op catalog

Effects listed per op are the semantic paths (02 §6) the op **declares**. G4 checks they changed exactly as predicted; G5 checks that nothing else changed.

#### Cell content

| Op | Fields | Declared effects |
|---|---|---|
| `set_value` | `target` (cell or range: broadcast), `value`, `number_format?`, `as_text?`, `rich_policy?: "reject" \| "replace" = "reject"` | `.value`, `.type` (+ `.rich` when replacing rich text; + `.style.number_format`) |
| `set_values` | `target` (range or top-left anchor), `values: Value[][]` | Same as above, per cell. The size must match the range |
| `set_formula` | `target`, `formula`, `kind?: "normal" \| "shared" \| "array"`, `ref?`, `cache?: "clear" \| "keep" \| {value}` | `.formula`, `.cache`, `wb.calc.fullCalcOnLoad` |
| `fill` | `target` (range), exactly one of `value` / `series: {start, step}` / `pattern_from: string` | Per-cell `.value`/`.formula` |
| `clear` | `target`, `what: ("values"\|"formats"\|"comments"\|"hyperlinks"\|"validation"\|"all")[] = ["values"]`, `remove_cells?: false` | The facets selected by `what` |
| `rich_set` | `target` (cell), `rich?: string` (markup) or `runs?: Run[]`, `phonetic?` | `.rich`, `.value` |
| `rich_style` | `target`, `at: Span`, `style: FontPatch` | `.rich` |
| `rich_replace` | `target`, `at: Span`, `text`, `style?: "inherit" \| FontPatch = "inherit"` | `.rich`, `.value` |
| `rich_insert` | `target`, `at: {before: Span} \| {after: Span} \| {index: number} \| "start" \| "end"`, `text`, `style?` | `.rich`, `.value` |
| `rich_delete` | `target`, `at: Span` | `.rich`, `.value` |
| `phonetic_set` | `target`, `runs`, `properties?` | `.phonetic` |

**Current `clear` slice:** only `what: ["values"]` (or omitted) is supported, for a cell or rectangular range of at most 500 cells. With `remove_cells: false` (default), existing cells retain their style; with `remove_cells: true`, existing `<c>` elements are removed while their rows remain. Absent cells remain absent in both modes. Clearing an existing formula invalidates its `calcChain` entry; a no-op clear leaves all package part bytes unchanged. Other facets fail closed.

**Current error-value slice:** `set_value` and `set_values` accept `{"error":"#N/A"}` and the canonical error tokens recognized by the legacy formula-cache writer. A new or existing cell stores `t="e"`; unsupported tokens fail before mutation. `excel_save.assert` accepts `equals.value: {"error":"#N/A"}` with exact readback.

**Current P2b slice:** `set_formula` accepts only `kind: "normal"` (or omitted) with `cache: "clear"` (default), `cache: "keep"` on an existing ordinary formula, or `cache: {"value": …}` with a finite number, string, boolean, or supported `{"error": "#N/A"}` token. Null, unsupported values, extra fields, and unsupported formula groups fail closed. An explicit cache writes a typed `<v>` for new or existing cells; its contents and type are checked on readback by G4 and independently on touched cells by G5. These gates verify the supplied value, not the formula's mathematical result; no recalculation occurs. Keeping the cache preserves its original result type and `<v>` content; G5 independently compares them and the element prefix against the source. With `cache: "clear"`, G4 rejects a leftover cache. Unsupported fields fail before mutating the session. Existing `set_value` formula objects remain compatible with P2a. `set_value` also broadcasts to a rectangular range (maximum 500 cells); existing per-cell rich-text and merged-cell guards still apply to the whole batch. `set_values` accepts a rectangular matrix (maximum 500 cells) with an exactly matching range or a top-left anchor, and expands to cell edits in one revision. `fill` accepts `value` or row-major exact decimal `series: {start, step}` on a rectangular range of at most 500 cells. Decimal/scientific operands are bounded to exponents −30 through 14; each result must have at most 15 integral digits and at most 15 significant digits. Out-of-range values, `pattern_from`, and additional fill fields fail closed. Bulk formatting/rich-text options are not supported yet; invalid shapes fail before applying any cells.

`Span = { match: string, occurrence?: number | "all" = 1, normalize?: "nfc" | "none" } | { range: [number, number] } | "all"`. Ranges count graphemes (02 §4.2).

Example: recolor one word in a three-run Vietnamese cell, guarded by the text the agent read.

```json
{ "op": "rich_style", "target": "A3",
  "expect": { "text": "Đỏ đậm nghiêng gạch chân 3" },
  "at": { "match": "đậm" }, "style": { "color": "FF0000FF" } }
```

#### Style

| Op | Fields | Declared effects |
|---|---|---|
| `style` | `target`, `patch: StylePatch` (02 §3.3), `apply_to_runs?: false` | `.style.<facet>.<field>` for each patched field (+ `.rich` when `apply_to_runs`) |
| `border` | `target` (range), `outline?: Side`, `inside?: Side`, `sides?: {left?, right?, top?, bottom?, diagonal?}` | `.style.border.*` on the affected edge cells only |
| `style_xf` (expert) | `target`, `xf: object`, `named_style?`, `exact_default_policy?` | `.style.*`, `xf_id` |
| `named_style_add` / `named_style_update` / `named_style_delete` | `name`, `style` / `patch` | `styles.named[Name]`; delete fails with `IN_USE` if cells still reference it |

#### Rows, columns and merges

| Op | Fields | Declared effects / transforms |
|---|---|---|
| `insert_rows` | `before: number` or `after: number`, `count = 1`, `format_from?: "above" \| "below" \| "none" \| number` | Transform; predicted rewrites of formulas, names, CF/DV ranges, tables, merges, hyperlinks, drawing anchors, chart series, print areas |
| `delete_rows` | `rows: string \| string[]` (e.g. `"5:7"`) | Transform; same rewrite set; references fully inside the deleted area become `#REF!` (listed explicitly) |
| `insert_cols` / `delete_cols` | `before`/`after: "C"`, `count` / `cols: "C:D"` | As for rows |
| `copy_rows` | `source: "3:4"`, `before?`/`after?`, `times = 1`, `include?: ("values"\|"formulas"\|"styles"\|"heights"\|"merges"\|"validations"\|"cf")[]` (default all) | Transform (insert) + content of the new rows. Replaces clone/copy/fill-rows |
| `copy_cols` | `source: "C"`, `before?`/`after?`, `times`, `include?` | As above |
| `merge` | `target` (range), `on_data_loss?: "error" \| "keep_origin_only" = "error"` | `S!merge`; with `keep_origin_only`, `.value` of the non-origin cells |
| `unmerge` | `target` | `S!merge` |
| `row_props` | `rows`, `height?: number \| null`, `hidden?`, `outline_level?`, `collapsed?`, `style?`, `phonetic?` | `S!row:N.<attr>`, **only the attributes given** (not the whole row, as in the legacy server) |
| `col_props` | `cols`, `width?: number \| null \| "autofit"`, `autofit?: {min, max}`, `hidden?`, `outline_level?`, `collapsed?`, `style?` | `S!col:L.<attr>` |

#### Sheets

| Op | Fields | Declared effects / transforms |
|---|---|---|
| `sheet_add` | `name`, `position?` or `after?` | `wb.sheets`, new parts, relationships, content types |
| `sheet_delete` | `name`, `on_references?: "error" \| "ref_error" = "error"` | `wb.sheets`, parts removed; with `ref_error`, every `#REF!` rewrite listed |
| `sheet_rename` | `name`, `new_name` | Transform: references in formulas, names, charts, data validation and CF rewritten |
| `sheet_move` | `name`, `position` | `wb.sheets` order, `localSheetId` of scoped names, active tab |
| `sheet_copy` | `name`, `new_name`, `position?` | New sheet plus its related parts |
| `sheet_copy_from` | `source_session`, `source_sheet`, `new_name?`, `position?` | New sheet; styles remapped into this workbook (new style records listed) |
| `sheet_state` | `name`, `state` | `wb.sheets[name].state`; refuses to hide the last visible sheet |
| `sheet_props` | `name`, `patch` (code name, tab color, filter mode, published, outline, page-setup props) | `S!props.*` |
| `sheet_views` | `name`, `views`, `mode: "replace" \| "patch"` | `S!view[i].*` |
| `freeze` | `name?`, `at: "B2" \| null` | `S!view[0].pane` |
| `sheet_protection` | `name`, `enabled?`, `patch?`, `password?` (hashed by the server) or `hash?` | `S!protection.*` |
| `protected_ranges` | `name`, `ranges`, `mode` | `S!protection.ranges` |
| `ignored_errors` | `name`, `rules`, `mode` | `S!ignored_errors` |

#### Data features

| Op | Fields | Declared effects |
|---|---|---|
| `table_add` / `table_update` / `table_delete` | `sheet`, `name`, `ref`, `columns?`, `header_row?`, `totals?`, `style?`, `autofilter?` / `name`, `patch` / `name`, `keep_data = true` | `S!table[Name].*`, table part, relationships |
| `autofilter` | `sheet`, `ref? \| null`, `columns?`, `sort?`, `mode: "patch" \| "replace"` | `S!autofilter`, `name[_xlnm._FilterDatabase@S]` |
| `cf_add` / `cf_update` / `cf_delete` | `sheet`, `sqref`, `rule` / `id`, `patch` / `id` | `S!cf[id].*`, dxfs |
| `dv_add` / `dv_update` / `dv_delete` | `sheet`, `sqref`, `rule` / `id`, `patch` / `id` | `S!dv[id].*` |
| `name_add` / `name_update` / `name_delete` | `name`, `value`, `scope?`, `hidden?`, `comment?` / `name`, `scope?`, `patch` / `name`, `scope?` | `name[Name@scope].*` |
| `hyperlink_set` / `hyperlink_remove` | `target`, `url?` or `location?`, `display?`, `tooltip?` / `target` | `.hyperlink`, relationships for external targets |
| `comment_set` / `comment_remove` | `target`, `text` or `rich`, `author?` / `target` | `.comment`, comments part, VML |

#### Printing

| Op | Fields | Declared effects |
|---|---|---|
| `print_area` | `sheet`, `areas: string[] \| null` | `name[_xlnm.Print_Area@S]` |
| `print_titles` | `sheet`, `rows?: "1:2" \| null`, `cols?: "A:A" \| null` | `name[_xlnm.Print_Titles@S]` (keeping attributes such as `hidden`, see EX-07) |
| `page_setup` / `print_options` | `sheet`, `patch`, `present?` | `S!print.page_setup.*` / `S!print.options.*` |
| `header_footer` | `sheet`, `sections`, `properties?` | `S!print.header_footer.*` |
| `page_breaks` | `sheet`, `rows?`, `cols?`, `mode` | `S!print.breaks` |

#### Workbook

| Op | Fields | Declared effects |
|---|---|---|
| `workbook_props` | `patch`, `date_system_policy?: "keep_serials" \| "shift_dates"` (required when changing `date1904`) | `wb.props.*` |
| `calc_props` | `patch` | `wb.calc.*` |
| `workbook_protection` | `enabled?`, `patch?`, `password?` or `hash?` | `wb.protection.*` |
| `workbook_views` | `views`, `mode` | `wb.views[i].*` |
| `doc_props` | `core?`, `app?`, `custom?: ({name, type, value} \| {name, delete: true})[]`, `modified?: "preserve" \| "now" = "preserve"` | `doc.*` |

#### Drawings

| Op | Fields | Declared effects |
|---|---|---|
| `image_add` | `sheet`, `anchor: "B2" \| "B2:F10"`, `source_path` or `base64`, `mime?`, `size?`, `name?`, `alt_text?` | `S!drawing[new]`, media part, relationships, content types |
| `chart_add` | `sheet`, `type`, `source`, `anchor`, `title?`, `options?` | `S!drawing[new]`, chart part, … |
| `shape_add` | `sheet`, `type`, `anchor`, `text?` or `rich?`, `style?`, `name?` | `S!drawing[new]` |
| `shape_text` / `shape_style` | `id: "dr_…"`, `text?` or `rich?` / `fill?`, `outline?`, `outline_width_pt?`, `text_color?` | `S!drawing[id].*` |
| `drawing_delete` | `id` | `S!drawing[id]`, orphaned parts removed |

#### Package (expert only, `expert: true`)

| Op | Fields | Declared effects |
|---|---|---|
| `part_upsert` | `part`, `content`, `encoding: "text" \| "base64"`, `content_type?` | `pkg:/part`, `ct:*` |
| `part_delete` | `part` | `pkg:/part`; fails if anything still references it |
| `rels_set` | `source_part`, `relationships` | `rel:/source#*` |
| `content_types_set` | `defaults?`, `overrides?` | `ct:*` |

## 5. `excel_save`

```ts
input {
  session: string,
  mode: "overwrite" | "save_as" | "copy",            // REQUIRED, no default (decided 2026-10-01)
       // overwrite: writes the session source; `path` omitted or equal to it.
       // save_as: `path` required and different; the session follows it.
       // copy: `path` required and different; the session stays on the source and remains dirty.
  path?: string,
  assert?: Assertion[],
  accept?: { id: string, reason: string }[],         // gap ids (g_…) of required gaps only
  override?: { id: string, reason: string }[],       // diff ids (d_…) of G5 `undeclared` differences only
  oracles?: { recalc?: "off" | "report" | "write_cache" = "off",
              render?: { targets: string[] } | false = false,
              open_check?: boolean = false },
  allow_signature_invalidation?: boolean = false,
  report?: "summary" | "full" = "summary"
}

Assertion = {
  target: string,                                    // A1 (sheet-qualified) or semantic path
  equals?: { value?: Value, display?: string, formula?: string, rich?: string, style?: StylePredicate },
  unchanged?: boolean,                               // compared against O after transforms
  except?: string[]
}

output data {
  status: "verified" | "verified_with_gaps" | "saved_with_overrides",
  accepted: { id: string, reason: string, gap: Gap }[],         // always listed, never hidden
  overrides: { id: string, reason: string, failure: Failure }[],
  path: string, sha256: string, size: number,
  backup?: { path: string, sha256: string, expires_at: string },
  revision_saved: number,
  gates: Gate[],
  assertions: { index: number, status: "verified" | "failed" | "unverified", expected?: any, actual?: any }[],
  readback: Record<string, Cell>,                    // re-read from the written file, not from the session
  oracles?: { recalc?: object, render?: { target: string, image_path: string }[], open_check?: object },
  verify_call: string                                // excel_verify(...) call that reproduces the check
}

Gate = { id: "G1".."G8", name: string, status: "verified" | "failed" | "unverified",
         checked: number, failures: Failure[], gaps: Gap[] }
Failure = { diff_id: string, path: string, expected: any, actual: any, kind: string, explanation?: string }
Gap     = { gap_id: string, path: string, reason: string, required: boolean }
```

- **Current P2b save-mode slice:** `mode` is required. `copy` requires a new path and leaves the session dirty; `save_as` requires a new path and follows it; `overwrite` uses the current session path (an optional `path` must agree). All three modes stage and run the same G1–G7 checks before publication. For followed saves, the session keeps a private baseline snapshot so subsequent edits and undo below `saved_revision` still replay from O; external changes to the followed file block further work. `overwrite` atomically replaces the destination and retains a local `.bak` of its prior bytes. Backup expiry/cleanup policy, report options, accept/override, optional oracles and full save-envelope metadata are not implemented yet.
- **Current P2b assertion slice:** `assert` accepts sheet-qualified cell `equals.value` (blank, string, boolean, supported error tokens, bounded numeric lexemes), `equals.formula`, and standalone `unchanged: true`. `unchanged` compares the source and staging cell XML subtrees, resolved values, and referenced shared-string `<si>` XML; absent cells must stay absent. It does not compare unrelated cells or package metadata (the other save gates handle those). No structural transforms are supported in this slice. Other assertion facets and combinations with `unchanged` fail closed; `equals.value` on a formula checks its cached result (including its type) without recalculating the formula. `null` requires a missing cache; unsupported cache types fail closed.
- **Blocked save:** `SAVE_BLOCKED` comes back with `details = { gates, assertions, blocked_by: Failure[] | Gap[] }` and the file is not written.
- **Gates** are defined in [04 §3](04-correctness.md#3-save-gates).
- **`verified_with_gaps`** is only possible when every remaining gap is optional (an oracle that is not available), or is a required gap accepted with a reason.
- **`saved_with_overrides`** means G5 found undeclared differences and the caller chose to write them anyway, each with a reason. It is **never** reported as verified, and agents must only use it with the user's explicit consent.
- **Not overridable:** failures of G1 (package integrity), G2 (new schema errors), G3 (namespaces, MC, prefix fidelity), G4 (intent), G6 (advanced parts) and G7 (the agent's own assertions). An `override` naming one of these fails with `NOT_OVERRIDABLE`. Rules: 04 §4.

## 6. Other tools

### 6.1 `excel_verify`

```ts
input {
  after_path: string,
  before_path?: string,            // default: newest unexpired backup of after_path, if any
  session?: string,                // when given, its ledger supplies effects and transforms → G4
  assert?: Assertion[],            // explicit expectations about after_path (same type as excel_save) → G7
  declared?: {                     // explicit intent for files produced elsewhere
    effects?: { path: string, after: any }[],                                   // → G4
    transforms?: { sheet: string, kind: "insert_rows" | "delete_rows" | "insert_cols" | "delete_cols",
                   at: string, count: number }[]                                 // → transform-aware G5
  },
  accept?: { id: string, reason: string }[],
  max_differences?: number = 200,
  detail?: "summary" | "full" | "package" = "summary"
}
output data {
  mode: "validate" | "compare",
  files: { before?: FileInfo, after: FileInfo },
  gates: Gate[],
  differences: { id: string, path: string, category: string, before: any, after: any,
                 classification: "declared" | "normalized" | "undeclared" | "unverified" }[],
  truncated: boolean
}
```

| Called with | Gates run | Replaces |
|---|---|---|
| `after_path` only, no backup found | G1–G3 | `excel_validate_workbook` |
| `after_path` + `before_path` | G1–G3 on both files + G5 (every difference counts as undeclared) + G6 | `excel_verify_preservation`, `excel_diff_package` |
| + `session`, or `declared.effects` | adds G4 | — |
| + `assert` | adds G7 | — |
| + `declared.transforms` | G5 maps O through the transforms before comparing | — |

**Current implementation (2026-10-03):** `after_path` alone runs partial G1/G3 and detached G2 on workbook, worksheet, and shared-string XML. `schema_issues` fail the call; `schema_gaps` list unsupported XML roots. With explicit `before_path`, both packages are checked and the decompressed bytes of every part are compared; ZIP-only differences are ignored. Added/removed/changed parts get stable hash IDs and `undeclared` classification, with `max_differences` limiting report entries (1–5000; default 200). This is **part-level G5 only**, not facet-level normalization or intentional-effect handling; an identical pair stays `unverified` unless `assert` (currently `equals.value`, `equals.formula`, `unchanged: true` with `before_path`) supplies an explicit expectation. A failed assertion returns `ASSERTION_FAILED` when no other difference is found, including the lost-edit case where the files are identical. Pre-existing schema errors are reported on both files but do not alone block compare; `new_schema_issues` block, and edits under a parent with a baseline schema error appear as required `schema_gaps` (masking). Package or markup errors still block before byte comparison. G4, G6, unsupported G7 facets and full G1–G3/G5 remain `unverified`. Neither file is changed.

**What a two-file comparison cannot see.** If an edit never reached the file (V-06), W equals O and there is no difference to report. A lost edit is only detectable when the expectation is supplied: through `session`, `declared.effects` or `assert` (e.g. `{ "target": "S!B2", "equals": { "value": "đã sửa" } }`). Likewise, after a row/column insert or delete, every shifted cell differs from O; losses among them are only separable when the transforms are known.

### 6.2 `excel_render`

```ts
input {
  session?: string, path?: string,                    // exactly one
  targets: string[],                                  // sheet-qualified ranges, or a sheet name
  provider?: "auto" | "excel" | "libreoffice" = "auto",
  scale?: number = 1, gridlines?: boolean = true, headers?: boolean = true,
  output_dir?: string
}
output: ImageContent blocks (PNG) + data {
  images: { target: string, path: string, width: number, height: number }[],
  provider: string, fidelity_notes: string[]
}
```

- With `session`, the current session state is written to a temp file through the normal writer (gates G1–G3 only) and rendered. Nothing is saved to the destination.
- `fidelity_notes` gives known LibreOffice-vs-Excel differences when LibreOffice is used.
- **Errors:** `ORACLE_UNAVAILABLE`, `TIMEOUT`.

### 6.3 `excel_export`

```ts
input  { session: string, what: "images" | "part" | "range_csv",
         sheet?: string, target?: string, part?: string,
         output_dir?: string, output_path?: string, overwrite?: boolean = false }
output data { files: { path: string, sha256: string, size: number, source: string }[] }
```

### 6.4 `excel_peek`

```ts
input  { path: string, detail?: "info" | "summary" | "preview" = "summary",
         sheet?: string, max_rows?: number = 20, max_cols?: number = 10 }
output data { sheets: {...}[], features?: {...}, preview?: { sheet: string, markdown: string }[] }
```

**Current P1 slice (not the full contract above):** `info` and `summary` expose sheet names, order, state, part path and `used_range` computed from explicit `<sheetData>/<row>/<c r>` coordinates (including empty styled cells), not from the possibly stale `<dimension>`. This does not include row-only formatting. `preview` adds top-left Markdown for the selected sheet (first sheet by default). Limits are 100 rows, 20 columns and 2,000 cells; responses include `used_range_basis: "explicit_cells"`, `partial: true` and `unverified: ["features"]`. No session is created.

### 6.5 `convert_to_markdown`

```ts
input  { file_path: string, sheet_name?: string, range_ref?: string,
         max_rows?: number, max_cols?: number, include_styles?: boolean = false }
output TextContent: Markdown (A1 column headers, Excel row numbers)
```

The legacy signature is kept so existing prompts keep working. The output format changes: A1 headers replace the legacy 0-based `col_n` and `row_index` annotations.

## 7. Error codes

| Code | Meaning | Retryable | Typical hint |
|---|---|---|---|
| `INVALID_INPUT` | Schema violation (pointer in `details`) | no | Fix the field |
| `INVALID_OP` | Unknown op, or invalid field combination | no | See the op catalog |
| `INVALID_ADDRESS` | Unparseable or out-of-range A1 | no | — |
| `FILE_NOT_FOUND`, `PATH_NOT_ALLOWED` | — | no | — |
| `UNSUPPORTED_FORMAT` | `.xls`/`.xlsb`/unknown | no | Convert to `.xlsx` first |
| `PACKAGE_INVALID` | Cannot open the package | no | `excel_verify` for details |
| `SESSION_NOT_FOUND`, `SESSION_BUSY` | — | busy: yes | Retry after the running operation |
| `READ_ONLY_SESSION` | apply/save on a read-only session | no | Reopen without `read_only` |
| `REVISION_CONFLICT` | `base_revision` ≠ current | yes | Re-read, then retry with the current revision |
| `PRECONDITION_FAILED` | An `expect` did not match | no | Re-read the target; `details` has expected/actual |
| `TARGET_NOT_FOUND`, `TARGET_AMBIGUOUS` | Target resolution failed | no | `details.candidates` |
| `TARGET_DELETED_IN_BATCH` | A later op targets something an earlier op deleted | no | Reorder, or use `sequential` mode |
| `MERGED_NON_ORIGIN` | Target is inside a merge but not its origin | no | `details.origin` |
| `RICH_TEXT_POLICY_REQUIRED` | `set_value` on a rich-text cell without a `rich_policy` | no | Use a `rich_*` op or `rich_policy: "replace"` |
| `AMBIGUOUS_FORMULA_TEXT` | `set_value` with a string starting with `=` | no | `set_formula`, or `as_text: true` |
| `MERGE_WOULD_DISCARD_DATA` | Merging over non-empty non-origin cells | no | `on_data_loss: "keep_origin_only"` |
| `IN_USE` | Deleting a style/name/part that is still referenced | no | `details.references` |
| `EXPERT_REQUIRED` | Expert op without `expert: true` | no | — |
| `UNSAVED_CHANGES` | Closing a dirty session | no | Save, or pass `discard_unsaved: true` |
| `SOURCE_CHANGED_ON_DISK` | The source changed since open | no | Close and reopen, or save with `save_as` |
| `DESTINATION_CHANGED` | The destination changed during save | yes | Retry |
| `SIGNATURE_WOULD_BREAK` | The save would invalidate digital signatures | no | `allow_signature_invalidation: true` |
| `SAVE_BLOCKED` | One or more gates failed or are unverified and not accepted | no | Read `details.gates` |
| `NOT_OVERRIDABLE` | `override` names a failure outside G5, or `accept` names a failure instead of a gap | no | Fix the cause; only G5 undeclared differences can be overridden |
| `BACKUP_FAILED` | The backup could not be created or verified | yes | — |
| `ORACLE_UNAVAILABLE` | The requested oracle is not installed | no | `excel_status` shows availability |
| `LIMIT_EXCEEDED` | Size/count limit | no | Paginate, or narrow the target |
| `CANCELLED`, `TIMEOUT` | — | yes | — |
| `INTERNAL_INTENT_MISMATCH` | The engine's own post-apply check failed: an engine bug | no | Report it. The session is unchanged |
| `INTERNAL_ERROR` | Unexpected exception | maybe | — |

## 8. Open design point: one big `excel_apply` schema

A single `oneOf` with about 75 op variants gives a large `inputSchema`. Plan:

1. Measure the serialized schema size in spike S8 and compare it with the legacy surface (112 tool schemas).
2. If it is too large, keep the transaction model but split the ops across three tools with identical semantics: `excel_apply` (cells, rich text, styles), `excel_apply_structure` (rows, cols, sheets), and `excel_apply_features` (tables, CF, DV, names, links, comments, print, views, protection, workbook, drawings). Package ops stay under `expert`.
3. Decide by measuring agent success rate in the eval suite (06 §2.7), not by guesswork.
