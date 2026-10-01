# 02 — Data model and conventions

**Principle: the agent issues commands and reads results in the same representation.** Whatever the agent sends in (addresses, values, styles, rich-text markup), results come back in that same form, so the agent can compare them directly instead of interpreting them.

## 1. Addressing

### 1.1 A1 strings (the only form for cells, ranges, rows and columns)

| Form | Example | Notes |
|---|---|---|
| Cell | `B5` | 1-based rows, letter columns, exactly as in Excel |
| Range | `A1:C10` | Both ends **inclusive** (unlike the legacy exclusive `end_row`) |
| Rows | `5:7`, `5` | Whole rows |
| Columns | `B:D`, `C` | Whole columns |
| Sheet-qualified | `Sheet1!B5`, `'Doanh thu Q3'!A1:C10` | Quote names containing spaces or special characters with `'…'`; a `'` inside a name is written `''` |
| Multi-range | `["A1:B2", "D4"]` | Only where OOXML allows it (conditional formats, data validation, print area) |

- Each tool or op has a default `sheet` field. A sheet-qualified address overrides it.
- There are no 0-based coordinates anywhere. The conversion table from legacy conventions is in [05 §2](05-legacy-mapping.md#2-conventions).
- Valid range is `A1`…`XFD1048576`. Anything outside fails with `INVALID_ADDRESS`.

### 1.2 Structured targets (when A1 is not used)

```jsonc
{ "name": "DoanhThu" }                                       // workbook-scoped defined name
{ "name": "DoanhThu", "scope": "Sheet1" }                    // sheet-scoped defined name
{ "table": "TblBan", "column": "Số lượng", "part": "data" }  // part: all|data|headers|totals
{ "match": { "text": "Tổng cộng", "in": "value" }, "within": "A:A", "occurrence": 1 }
```

- A `match` target must resolve to **exactly one** address. No hit → `TARGET_NOT_FOUND`. More than one without `occurrence` → `TARGET_AMBIGUOUS`, with the candidates listed.
- Results always include the resolved A1 address, so the agent knows where the op landed.

### 1.3 Merged cells

- Targeting a cell of a merged range that is not its top-left origin fails with `MERGED_NON_ORIGIN`, and `details.origin` names the origin cell. The legacy server calls these "slave cells".
- Reading a non-origin cell returns `merged: { "origin": "B2", "range": "B2:D3" }`.

## 2. Values

### 2.1 Types

| `type` | JSON input | Stored in OOXML | Notes |
|---|---|---|---|
| `blank` | `null` | Cell without `<v>`, or no cell at all | `clear` distinguishes "remove the cell" from "keep an empty styled cell" |
| `number` | `123.45` | `<v>` | At most 15 significant digits are stored. Longer numbers raise warning `PRECISION_LOSS` |
| `string` | `"0123"` | sharedStrings or inlineStr | **No coercion**: `"0123"` stays a string |
| `bool` | `true` | `t="b"` | — |
| `error` | `{"error": "#N/A"}` | `t="e"` | — |
| `date` | `{"date": "2026-10-15"}`, `{"datetime": "2026-10-15T08:30:00"}`, `{"time": "08:30"}` | Serial number according to `date1904` | If the cell has no date format, the op must include `number_format`; otherwise warning `DATE_WITHOUT_FORMAT` |

- A string starting with `=` passed to `set_value` fails with `AMBIGUOUS_FORMULA_TEXT`. Use `set_formula` for a formula, or add `"as_text": true` to store the literal text. The agent can never accidentally turn a formula into text, or text into a formula.
- String storage (sharedStrings vs inlineStr): **new** cells use sharedStrings, as Excel does. **Existing** cells keep their current storage.

### 2.2 Display value (`display`)

This is the string the user sees in the cell.

- **How it is computed:** from the value, the numFmt code, `date1904`, and Excel's General formatting rules.
- **Display locale (decided: `en-US`):** `display` is rendered for `DOCLOUPE_EXCEL_DISPLAY_LOCALE`, default `en-US`, and every read reports the locale used (`display_locale`). Excel itself renders with the viewer's regional settings, in three ways:
  - the thousands and decimal separators of **every** numeric format (`#,##0.00` shows `1,250,000.00` under en-US but `1.250.000,00` under vi-VN);
  - built-in date/time formats (e.g. id 14 `m/d/yyyy`);
  - month/day names without an explicit `[$-xxx]` locale tag.

  Cells affected by any of these come back with `locale_dependent: true`, so the agent knows that a user on another locale sees a different string.
- **No OS culture data:** locale rendering (separators, month/day names, currency symbols) uses the server's own locale tables, not .NET/OS culture data. The server runs with `InvariantGlobalization=true` (01 §12), so `display` is identical on every machine.
- **Formula cells:** `display` comes from the cached value. A missing or stale cache gives `display: null` with `cache: "missing" | "stale"`.
- **Screen-only effects** (truncation, `####` in a narrow column, wrapping) are **not** derived from data. They require `excel_render` (see 04 §7).

## 3. Effective style

### 3.1 Facets

```jsonc
{
  "font":      { "name": "Calibri", "size": 11, "bold": true, "italic": false, "underline": "single",
                 "strike": false, "color": "FFC00000", "vert_align": "superscript",
                 "family": 2, "charset": 163, "scheme": "minor",
                 "outline": false, "shadow": false, "condense": false, "extend": false },
  "fill":      { "pattern": "solid", "fg": "FFFFFF00", "bg": "auto" },      // or { "gradient": {...} }
  "border":    { "left": {"style": "thin", "color": "FF000000"}, "right": …, "top": …, "bottom": …,
                 "diagonal": {"style": "thin", "up": true, "down": false} },
  "alignment": { "horizontal": "center", "vertical": "top", "wrap": true, "shrink": false,
                 "indent": 0, "rotation": 0, "reading_order": "context" },
  "number_format": "#,##0.00 \"₫\"",
  "protection": { "locked": true, "hidden": false },
  "named_style": "Normal"
}
```

### 3.2 Output

- **Sparse by default:** only facets that differ from the workbook's default cell style (`cellXfs[0]`) are returned. This saves tokens, and the agent sees only what actually differs.
- **Full view:** `style_detail: "full"` returns every facet plus `xf_id` and the raw xf data, for experienced callers.
- **Colors are always resolved to ARGB**, with their source when there is one: `"color": {"rgb": "FFC00000", "theme": 5, "tint": -0.25}`. Theme colors come from `theme1.xml`; indexed colors come from the default 64-color palette or the workbook's own `<colors>`. The agent thinks "dark red"; the file stores `theme=5 tint=-0.25`; the result must carry both.

### 3.3 Input (patch semantics)

- **How keys are interpreted:**
  - key present → set the value;
  - `null` → unset (back to default / inherited);
  - key absent → **unchanged**.
- Explicit `false` and `0` are preserved as such, distinct from "not set". The legacy server follows the same rule.
- **Accepted color forms:** `"C00000"`, `"FFC00000"`, `{"theme": 5, "tint": -0.25}`, `{"indexed": 10}`, `"auto"`.
- **Style records:** the engine reuses identical existing font/fill/border/xf records, adds new records only when needed, and **never renumbers** existing styles.

### 3.4 Rich text vs cell font

- A run without `rPr` uses the cell font.
- A run with `rPr` takes the properties present in `rPr` and falls back to the cell font for the rest. **This merge rule must be confirmed against the Excel oracle (spike S7) before it is final.**
- A cell-level `style` op does **not** disturb runs. To push a font change into the runs, pass `"apply_to_runs": true` (see the `style` op in 03).

## 4. Rich text

### 4.1 Markup

The markup is the default way to read and write rich text:

```
<r b color="FF0000">Đỏ đậm </r><r i sz="14">nghiêng </r><r u font="Times New Roman">gạch chân 1</r>
```

**Grammar:**

```ebnf
rich      = { plain | run } ;
run       = "<r" { ws attr } ">" text "</r>" ;
plain     = text ;                          (* run without rPr: uses the cell font *)
attr      = name [ "=" '"' value '"' ] ;    (* no value means true *)
text      = { char | "&lt;" | "&gt;" | "&amp;" | "&quot;" } ;
```

| Attribute | Values | OOXML element |
|---|---|---|
| `b`, `i`, `s` (strike), `outline`, `shadow`, `condense`, `extend` | no value = true; `="0"` = **explicit false** | `<b/>`, `<i/>`, `<strike/>`… |
| `u` | no value = `single`; `double`, `singleAccounting`, `doubleAccounting` | `<u val=…/>` |
| `sz` | number | `<sz val=…/>` |
| `font` | font name | `<rFont val=…/>` |
| `color` | `FF0000`, `FFFF0000`, `theme:5`, `theme:5,-0.25`, `indexed:10`, `auto` | `<color …/>` |
| `va` | `superscript`, `subscript`, `baseline` | `<vertAlign/>` |
| `family`, `charset` | number | `<family/>`, `<charset/>` |
| `scheme` | `minor`, `major`, `none` | `<scheme/>` |

- **Whitespace** and line breaks are kept **verbatim**. In JSON a line break is `\n`. The writer adds `xml:space="preserve"` when needed; the agent never deals with it.
- **Phonetics** are not part of the markup. They have their own field: `"phonetic": {"runs": [{"text": "…", "range": [0, 2]}], "properties": {"type": "fullwidthKatakana", "alignment": "left", "font_id": 0}}`.
- **Structured form for code:** `runs: [{ "text": "…", "font": {…} }]`. The two forms convert 1:1, and a property test requires `parse(render(runs)) == runs`.

### 4.2 Unicode, graphemes and Vietnamese

- **Content is stored verbatim.** The file's text is never auto-normalized to NFC or NFD.
- **Matching:** `match` compares after NFC-normalizing **both sides** (default `normalize: "nfc"`; `"none"` for exact matching), then maps the hit back to grapheme boundaries in the original text.
- **Offsets** (`range: [start, end]`) count **graphemes**, not code points or UTF-16 units. No op may split a grapheme.
  - Example: NFD "ệ" (`e` + U+0323 + U+0302) is one grapheme.
  - The legacy defect V-05 is exactly this: it split the text into `"Vie"` and `"̣̂t"`.
- **Case-insensitive** matching uses `InvariantCulture` with `IgnoreCase`, which handles `Đ`/`đ` correctly.
- **Text form:** in the `full` read view, each string cell carries `text_form: "NFC" | "NFD" | "mixed"`, so the agent knows what it is dealing with.

## 5. Formulas

| Aspect | Read | Write |
|---|---|---|
| Text | `"formula": "=SUM(B2:B10)"`: always with a leading `=`, en-US syntax as stored in OOXML | Accepted with or without `=` |
| Kind | `kind`: `normal` \| `shared` \| `array` \| `dataTable` | `set_formula.kind` |
| Shared | Each cell returns **its own expanded** formula, plus `shared: {"si": 3, "master": "B2", "ref": "B2:B20"}` | Editing one member detaches it from the shared group (a declared effect) |
| Array | `array: {"ref": "C2:C5"}` | — |
| Dynamic array / spill | `spill: {"ref": …}`; the `cm` attribute and `xl/metadata.xml` are preserved | Creation not supported in the first phases |
| Cache | `value` comes from the cache; `cache: "fresh" \| "stale" \| "missing" \| "unknown"` | `cache: "clear"` (default) \| `"keep"` \| `{ "value": … }` |

- **`stale`** means an op in this session changed an input the formula references directly, on the same sheet or another sheet.
- **`unknown`** is used for formulas with `INDIRECT`, `OFFSET`, or external references, where dependencies cannot be derived.
- **`fullCalcOnLoad`:** any op that changes a value or a formula sets `calcPr fullCalcOnLoad="1"`. This is a declared effect, not a hidden side effect, and it makes Excel recalculate on open.

## 6. Semantic paths

This is the shared language for effects, diffs, `accept`, and `assert`.

```
path        = wb-path | doc-path | sheet-path | name-path | pkg-path
wb-path     = "wb." facet                 ; wb.calc.fullCalcOnLoad, wb.props.date1904, wb.protection, wb.views[0]
doc-path    = "doc." ("core"|"app") "." field | "doc.custom[" Name "]"
sheet-path  = sheet "!" target [ "." facet { "." facet } ]
sheet       = SheetName | "'" QuotedName "'"
target      = A1 | "row:" N | "col:" Letters | "merge" | "cf[" id "]" | "dv[" id "]"
            | "table[" Name "]" | "print" | "view[" i "]" | "protection" | "drawing[" id "]"
            | "autofilter" | "ignored_errors" | "props"
name-path   = "name[" Name [ "@" sheet ] "]"
pkg-path    = "pkg:/" PartName | "rel:/" SourcePart "#" Id | "ct:default:" Ext | "ct:override:/" PartName
```

**Examples:**

- `S!B5.value`
- `S!A3.rich`
- `S!A3.style.font.color`
- `S!row:5.height`
- `S!col:C.width`
- `'Doanh thu'!merge`
- `S!cf[cf_8f2a].rule.formula[0]`
- `name[_xlnm.Print_Area@S]`
- `pkg:/xl/charts/chart1.xml`
- `rel:/xl/workbook.xml#rId4`

**Cell facets:** `value`, `type`, `formula`, `formula.kind`, `cache`, `rich`, `phonetic`, `style.<facet>.<field>`, `hyperlink`, `comment`, `present` (whether the cell element exists in the XML).

**Diff ids:** `d_` + the first 10 hex chars of `sha256(path ‖ before ‖ after)`. Diff ids are stable across runs on the same data, so they can be used in `accept` (see 04 §4).

## 7. Hashes and identifiers

| Kind | How it is computed | Used for |
|---|---|---|
| Cell `hash` | First 12 hex chars of sha256 over canonical JSON of value, type, formula, rich, phonetic, effective style, hyperlink, comment | `expect: {"hash": …}` guarantees the agent edits exactly what it read |
| Range `hash` | sha256 over the ordered cell hashes plus merge info | `expect` for a whole range |
| `revision` | Session-wide integer | `base_revision` detects overlapping edits |
| Object ids | Conditional formats `cf_xxxx`, validations `dv_xxxx`, drawings `dr_xxxx`: derived from content at load. Tables by name; defined names by `Name@scope` | Replace the legacy 1-based `shape_index` and per-session rule ids. An id stays stable across revisions until that object itself is edited |

## 8. Cell representation when reading

```jsonc
{
  "addr": "A3",
  "type": "string",
  "value": "Đỏ đậm nghiêng gạch chân 3",
  "display": "Đỏ đậm nghiêng gạch chân 3",
  "rich": "<r b color=\"FF0000\">Đỏ đậm </r><r i sz=\"14\">nghiêng </r><r u font=\"Times New Roman\">gạch chân 3</r>",
  "style": { "alignment": { "wrap": true } },          // sparse: only what differs from the default
  "hash": "9f2c41d0be7a"
}
{ "addr": "D3", "type": "number", "formula": "=C3*2", "value": 9, "display": "9", "cache": "fresh", "hash": "…" }
{ "addr": "C4", "merged": { "origin": "B4", "range": "B4:D4" } }
```

**Markdown view** (see `excel_read` in 03):

- column headers are the letters A, B, C…, and the first column holds Excel row numbers;
- rich-text cells are marked `†` and formula cells `ƒ`;
- read those cells again in the `cells` view to see their details.
