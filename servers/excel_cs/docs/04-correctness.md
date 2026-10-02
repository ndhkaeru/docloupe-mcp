# 04 — Correctness contract with the agent

## 1. The three questions

After every edit the agent must be able to answer, from the tool's output alone:

| # | Question | Proven by | Legacy server |
|---|---|---|---|
| Q1 | **Did my edit land exactly as I intended?** | G4 (intent): independent re-read of W, compared with the predicted after-values | ❌ Not checked. V-06 reports success while the edit is lost |
| Q2 | **Did anything else change?** | G5 (preservation) with transforms, plus G1–G3 and G6 | ⚠️ Partially: off by default, wildcard exemptions, blind on structural edits (V-04) |
| Q3 | **Will the user see what I think they will see?** | Display strings and effective style in the readback; G8 oracles (render, recalc, open check) | ❌ No display values; `excel_capture` exists but is not tied to edits |

## 2. Invariants

| ID | Invariant | Enforced by |
|---|---|---|
| I1 | Nothing in the written file differs from O except the declared effects (after transforms and registered normalizations). The only exception is G5 differences explicitly overridden with a reason, and the save then reports `saved_with_overrides` | G5, G6, 04 §4 |
| I2 | Every declared effect is present in the written file with exactly its predicted value | G4 |
| I3 | Every "verified" verdict comes from re-reading the **written file** with a reader that shares no code with the writer | Architecture rule (01 §3) + G4/G5 implementation |
| I4 | A check that did not run, or cannot model something, is `unverified`, never `verified` | Gate status rules (§3.9) |
| I5 | A batch of ops applies completely or not at all | `excel_apply` pipeline (01 §5.2) |
| I6 | Results use the same representation as the request (A1, typed values, markup, effective style, display) | Output DTOs (02, 03) |
| I7 | The agent can bind every op to what it read (`expect.hash`, `base_revision`) | Preconditions |
| I8 | The source file is never modified except by a successful `overwrite` save, which is preceded by a verified backup | Save pipeline (01 §5.3) |

## 3. Save gates

All gates run on the staging file before commit. Gates G1–G7 are **required**. G8 is optional and only runs when requested.

### 3.1 G1 — Package integrity

- The ZIP is valid, with no duplicate entries, and is within the limits of 01 §7.
- Required parts are present: `[Content_Types].xml`, `_rels/.rels`, the main workbook part **resolved through the `officeDocument` relationship in `_rels/.rels`** (usually `xl/workbook.xml`, but not necessarily), and that part's relationships part.
- Part names follow OPC rules: targets are resolved relative to their source part, percent-decoded, and looked up case-insensitively.
- Every internal relationship target exists. No part loses its last incoming relationship unless that removal was declared.
- Every relationship reference **inside a touched part** (`r:id`, `r:embed`, `r:link`, `r:pict`…) resolves to an existing relationship of that part. Detached schema validation does not check this (P2a SDK-role review finding 2).
- Every part has an effective content type, and the main part's content type matches the extension (`.xlsm` must be macroEnabled).
- *Legacy:* `inspect_xlsx_package` / `_validate_package_xml` check ZIP and XML well-formedness only.

### 3.2 G2 — Schema validity

- A separate `Schema` adapter runs `OpenXmlValidator` (`FileFormatVersions.Microsoft365`) on **detached typed roots** built from each touched/added part in staging, including semantic root attributes and the namespace bindings needed for validation, without calling `SpreadsheetDocument.Open` or allowing SDK serialization to write output. The source part in O is validated in the same way. Exact default-namespace declarations are checked by G3 on the raw XML; the SDK detached root cannot reproduce them lexically.
- The error set is compared by stable location and diagnostic with the baseline errors of the same parts in O; matching only error IDs or counts is insufficient. **New errors block the save**; pre-existing errors are reported as `baseline_errors` but do not block. A part for which detached validation cannot establish coverage is a required `unverified` gap, never `verified` (see §4). The P2a SDK-role spike established only first-worksheet detached validation on eight fixtures, with injected errors at three depths, not full G2.
- **Masking rule.** The validator reports only the first content-model error per parent element, so a baseline error hides any new error under the same parent. In the P2a SDK-role spike, fixture 01's baseline error on `<worksheet>` hid a new misplaced child of `<worksheet>`, while new errors under `<sheetData>` and inside a cell were still reported. Therefore:
  - an edit that changes the child list of an element E (inserting, removing or reordering children of E) is covered by G2 only if E has **no** baseline content-model error in O;
  - otherwise the edit is a required `unverified` gap (`reason: "g2_masked_by_baseline_error"`) and is not reported as covered.
  - For `set_value`, E is the edited `<c>`, or the `<row>` that receives a new cell.
- **Not covered by G2:** relationship references (`r:id` and similar) are not checked by detached validation and belong to G1 (§3.1). G2 is also the safety net for element order in hand-written System.Xml edits, but the engine must insert children in schema order to begin with.
- *Legacy:* no schema validation.

### 3.3 G3 — Markup compatibility and namespaces

- Every prefix listed in `mc:Ignorable`, `mc:ProcessContent` and `mc:MustUnderstand` is declared in scope. This catches EX-04.
- Every element and attribute keeps its namespace URI. For example, `x14ac:dyDescent` must not become a bare `dyDescent` (EX-03).
- **Lexical fidelity of touched parts.** The C# writer keeps what the source part looked like (decided 2026-10-01):
  - every namespace declaration on O's part root is present, **with the same prefix**, including the default namespace;
  - every element and attribute uses the same prefix as in O;
  - the XML declaration is identical to O's (present or absent, same `encoding`/`standalone`).

  A difference fails with `kind: "prefix_rewritten"` or `"declaration_changed"`.

  This is stricter than semantic equivalence on purpose. Downstream tools that read OOXML lexically break when a default namespace becomes `x:`: the legacy server's regex patch is one example (V-06). Spike S2 showed that Open XML SDK serialization makes exactly this change, so P2a uses the source-byte-preserving `System.Xml` subtree splice demonstrated on one cell in S2b, not a whole-sheet SDK serialization.
- *Legacy:* `_markup_compatibility_errors` exists but missed EX-04; the root namespace context is not captured.

### 3.4 G4 — Intent

- **What is compared:** for every declared effect `(path, before, after_expected)`, the value read from W is compared with `after_expected` after normalization (§5).
- **Status:**
  - `failed` when the value does not match: `kind: "intent_mismatch"`, with the expected value, the actual value, and an explanation when the cause is known.
  - `failed` when an effect is declared but the path is absent from W (`kind: "intent_missing"`).
  - Effects whose before equals after (no-ops) are reported as `info`.
- **Readback:** the `readback` block of the save report is produced by this gate's reader, so the agent sees exactly what is in the file.
- *Legacy:* not implemented. In V-06, both sides of the legacy comparison were identical, so it passed.

### 3.5 G5 — Preservation

- **What is compared:** for every facet and part that is **not** covered by a declared effect, `W == T(O)`, where `T` applies the batch transforms (§6).
- **Precision:** effects are fine-grained (`S!row:5.height`, not `S!row:5`, and never `S!*`). There are no wildcards.
- **Unmodeled content:** anything the snapshot does not model is compared as canonical XML (01 §5.4), so nothing goes uncompared.
- **Differences** are classified as `declared` (belongs to G4), `normalized` (a registered rule matched, with its id), `undeclared` (fails), or `unverified` (the verifier cannot model the change; a required gap).
- *Legacy:* the `_Changes` classifier with `REQUESTED` / `UNAPPROVED_LOSS`… Requested paths were derived as broad patterns (`worksheets/S/cells/B5*`, `worksheets/S/*` for structure), with no transform-aware comparison.

### 3.6 G6 — Advanced parts

- **VBA:** if O has `vbaProject.bin`, then W has the part, its relationship, and the macroEnabled content type (EX-01).
- **Digital signatures:** any change to a signed package invalidates the signatures. Without `allow_signature_invalidation` the gate fails (`SIGNATURE_WOULD_BREAK`). When allowed, it is reported as a declared effect.
- **Other advanced parts must be present and semantically equal unless declared:** charts, media, drawings, printer settings, pivot caches/tables, slicers, timelines, external links, `customXml`, `customUI`, threaded comments/persons, `xl/model`, `xl/metadata.xml`. This catches EX-01.
- **calcChain:** the writer removes `xl/calcChain.xml` only when an existing formula is replaced with non-formula content, and declares that removal as an effect. Otherwise the part must be unchanged. Value changes still set `fullCalcOnLoad` independently. Broader shared/array formula handling is deferred to P3.

**Current P2b subset:** the independent Verify gate fails closed on signed sources, compares every non-core part byte-for-byte (stronger than semantic equivalence), checks unchanged relationship and effective content-type mappings, and allows only the addition of shared strings or the justified removal of `calcChain`. Deliberate advanced-part edits, semantic normalizations, and signature-invalidation consent are not supported yet.

### 3.7 G7 — Agent assertions

- Each `assert` item in `excel_save` is evaluated on W:
  - `equals` compares specific facets;
  - `unchanged` compares against `T(O)`;
  - `except` lists paths to exclude.
- One failed assertion fails the gate.

### 3.8 G8 — Oracles (optional)

- **Gates:** `recalc`, `render`, `open_check`; see §7.
- **When unavailable:** a requested but unavailable oracle is reported as an *optional* gap. It yields `verified_with_gaps`, not a block.
- **Evidence only:** oracle results never turn a failed gate into a passed one.

### 3.9 Status rules

```
overall = blocked               if G1, G2, G3, G4, G6 or G7 has any failure   (never overridable)
        = blocked               if G5 has a failure that is not in `override`
        = blocked               if a required gate has an unverified gap that is not in `accept`
        = saved_with_overrides  if any G5 failure was overridden          (never "verified")
        = verified_with_gaps    if only optional gaps remain, or required gaps accepted with a reason
        = verified              otherwise
```

### 3.10 Comparison with the legacy classification

| Legacy | New | Change |
|---|---|---|
| `REQUESTED` | `declared` | Stricter: the value must also equal the predicted after-value (G4) |
| — | `intent_mismatch` / `intent_missing` | New: fails |
| `APPROVED_NORMALIZATION` | `normalized` (rule id) | Rules live in code and tests, not caller input |
| `UNAPPROVED_LOSS` | `undeclared` | Fails. Can only be written via `override`, which yields `saved_with_overrides`, never `verified` |
| `VERIFIER_GAP` | `unverified` (required gap) | Blocks unless accepted by gap id with a reason |
| `FIXTURE_GAP` | — | Test-harness concept only, moved to the evidence CLI |
| `PACKAGE_INVALID` | G1/G2/G3 failure | Split by cause |

## 4. Accepting gaps and overriding differences

The guiding rule: **a failure can never be turned into "verified".** Only what the verifier *could not check* can be accepted. A small class of *detected* differences can be written on purpose, but the result says so permanently.

| Mechanism | Takes | Allowed for | Requires | Result |
|---|---|---|---|---|
| `accept` | Gap ids (`g_…`) | Required gaps: something the verifier cannot model, e.g. `unmodeled_reference` | A `reason` per item | `verified_with_gaps`; every accepted gap is listed with its reason |
| `override` | Diff ids (`d_…`) | **Only** G5 `undeclared` differences | A `reason` per item, and the user's explicit consent (stated in the server instructions) | `saved_with_overrides`, **never** `verified`; listed in the report and recorded in the ledger |
| — | — | Failures of G1 (package integrity), G2 (new schema errors), G3 (namespaces, MC, prefix fidelity), G4 (intent mismatch or missing), G6 (advanced-part loss), G7 (the agent's own assertions) | — | Always `blocked`. Fix the cause. A G7 failure means the edit or the assertion is wrong |

- **Exact ids only.** `accept` and `override` take exact ids from a previous blocked report, never path patterns. Ids are content hashes of path + before + after (02 §6); if the difference changes, the old id no longer matches and the save blocks again.
- **Signatures** are not an override: invalidating a digital signature is governed by `allow_signature_invalidation` and declared as an effect (G6).
- **Visibility:** accepted gaps and overrides appear in the save report with their full details and reasons, so they cannot disappear from view.
- *Legacy:* `approved_normalizations` took caller-supplied rules (path-only entries became `VERIFIER_GAP`), and `requested_paths` took glob patterns. Both let a broad pattern hide unrelated losses.

## 5. Normalization registry

Rules are implemented in `Verify` and each one has a bidirectional fixture test (A→B and B→A both classified as `normalized` with the same rule id). A rule applies only within its stated scope.

| ID | Equivalence | Scope / condition | Example |
|---|---|---|---|
| N-BOOL | `<b/>` ≡ `<b val="1"/>` ≡ `<b val="true"/>`; `0` ≡ `false` | Boolean-typed attributes and element-presence booleans, per schema type table | V-02: `<b val="1"/>` vs `<b/>` |
| N-NUM | Lexical forms of the same `xsd:double`/`decimal` value: `14` ≡ `14.0` ≡ `1.4E1` | Numeric-typed attributes per schema (`sz`, `ht`, `width`…), compared as parsed numbers | V-02 `sz`, V-07 `ht="30"` vs `"30.0"` |
| N-ORDER | Child order is irrelevant | **Only** for content models that are unordered in the schema (e.g. run properties). Generated from the schema, not hand-listed | V-02: `<i/><sz/>` vs `<sz/><i/>` |
| N-ATTR | Attribute order is irrelevant | All XML | — |
| N-PREFIX | Different prefixes for the same namespace URI | **Only in `excel_verify` compare mode, for files written by other tools.** It never applies to parts written by the C# server, where G3 requires the original prefixes. MC tokens must still resolve | — |
| N-XMLSPACE | `xml:space="preserve"` present or absent | Only when the text has no leading/trailing whitespace and no line breaks | EX-06: D3 gained `xml:space` |
| N-DEFAULT | An attribute that is absent ≡ present with its schema default | Only when the value equals the schema default from the generated table | EX-08: `showButton="1"`, `hiddenButton="0"` |
| N-COUNT | `count` / `uniqueCount` attributes | Derived attributes, compared against the actual collection size instead | — |
| N-DIMENSION | `<dimension ref>` | Derived; must equal the actual used range of W | — |
| N-SST | Shared-string index vs text | Cells compare by text, runs and phonetics, not by index. sharedStrings vs inlineStr storage is equivalent **only in `excel_verify` compare mode**; the C# writer never changes the storage of untouched cells | — |
| N-STYLEIDX | Different `s` index | Equivalent when the resolved effective style is identical | — |
| N-RID | Relationship ids renumbered | Equivalent when the multiset of (source, type, target, mode) is equal **and** every referencing attribute is rewritten consistently | EX-08: table relationship id changed |
| N-CT | Default vs Override content-type entries | Equivalent when the effective content type of every part is equal | — |
| N-ZIP | Compression level, timestamps, entry order | ZIP container only | — |

**Not a normalization (must be declared):**

- removing `calcChain`;
- setting `fullCalcOnLoad`;
- detaching a cell from a shared formula;
- `#REF!` rewrites;
- invalidating signatures;
- remapping styles in `sheet_copy_from`.

## 6. Structural edits

### 6.1 Transform-aware comparison

A transform `T` maps O coordinates to W coordinates, for example `insert_rows(before=5, count=2)`: rows ≥ 5 move by +2. Cells in a deleted area map to ⊥.

- **Moved cells:** for every cell `c` of O, `W[T(c)] == O[c]`, except for facets with declared rewrites (formula references).
- **Deleted cells:** cells that map to ⊥ must be declared as deleted.
- **Inserted cells** must equal their predicted content (`format_from`, `copy_rows`).
- **Ranged objects** are compared after mapping their ranges through `T`:
  - merges, CF/DV sqref, tables and their autofilter, defined names, print area and titles, hyperlink refs;
  - drawing anchors, chart series formulas, sparkline refs (in `extLst`), pivot cache source ranges.
- **Legacy contrast:** the legacy server marked the whole sheet as requested (`worksheets/S/*`). In our probe, 204 changes were all classified `REQUESTED`, so a loss of rich-text runs in a shifted row would have passed unnoticed (V-04).

### 6.2 Independence of reference shifting

If the engine shifts a reference wrongly **and** declares that wrong prediction, G4 and G5 would both pass. To avoid this common-mode failure:

- `Verify` has its **own minimal reference shifter**, implemented separately from `Formula`. It covers A1/R1C1 refs, ranges, whole rows/columns, 3D refs and defined-name refs.
- G5 recomputes the expected rewrite with it and compares the result with the engine's declaration. A disagreement fails with `kind: "rewrite_disagreement"`.
- Where the verifier's shifter cannot model a construct, the result is a required gap (`g_…`, `reason: "unmodeled_reference"`), never a silent pass.
- The two implementations are differential-fuzzed against each other (06 §2.4).

### 6.3 Excel-faithful behavior that still needs a warning

- **`INDIRECT("B"&n)` and other text-built references** are not adjusted. Excel does not adjust them either, so the unchanged text is correct. The result still carries warning `TEXT_REFERENCE_NOT_SHIFTED` with the affected cells, because the agent may expect otherwise.
- **References into a deleted area** become `#REF!`, and each one is listed in the effects.

## 7. Oracles

| Oracle | Provider | What it adds | Limits |
|---|---|---|---|
| `recalc` | Excel COM (`CalculateFull`) preferred, LibreOffice otherwise | Fresh formula values for touched and dependent formulas, compared with the cached values; reported per formula | `write_cache` is allowed **only** with Excel; LibreOffice results can differ for some functions |
| `render` | Excel COM (`CopyPicture` → PNG exported by Excel), LibreOffice (PNG export of a temp copy whose print area is the target range). No rasterizer in the server (01 §12) | PNG of the target ranges, returned as MCP image content for multimodal agents | LibreOffice fidelity notes (fonts, conditional-format icons, sparklines) |
| `open_check` | Excel COM | "Excel opens the file without a repair prompt" | Detection method to be settled in spike S5. LibreOffice is too tolerant to count as evidence |

- Oracles open a **copy** of the staging file, with macros disabled and external links never updated.
- Oracle results are evidence attached to the report. They never replace G1–G7.

## 8. Agent workflow patterns

### 8.1 Default flow

```
excel_open
excel_read  view=cells target=…            → note hash / text
excel_apply base_revision=r ops=[{…, expect:{hash:…}}]
            → check results[].resolved, diff, readback
excel_save  mode=overwrite|save_as|copy (assert … optional)
            → status must be "verified"; check readback
```

### 8.2 Structural changes: preview first

```json
{ "session": "xs_…", "base_revision": 4, "dry_run": true,
  "ops": [ { "op": "insert_rows", "sheet": "Data", "before": 12, "count": 3, "format_from": "above" } ] }
```

The plan lists the transform, every formula/name/CF/table rewrite, and moved objects. The agent applies it only if the plan matches its expectations.

### 8.3 Assertions on save

```json
{ "assert": [
    { "target": "Data!E20", "equals": { "display": "1,250,000 ₫" } },
    { "target": "Data!A1:H11", "unchanged": true },
    { "target": "Data!A3", "equals": { "rich": "<r b color=\"FF0000\">Đỏ </r><r b color=\"0000FF\">đậm</r>…" } } ] }
```

### 8.4 Handling `SAVE_BLOCKED`

1. Read `error.details.blocked_by`. Each entry has a gate, a path, expected and actual values, and possibly an explanation.
2. When it is an engine/verifier disagreement (`intent_mismatch`, `rewrite_disagreement`), do **not** retry blindly. Report it.
3. When it is a required gap for something the verifier cannot model, the agent may retry `excel_save` with `accept: [{ "id": "g_…", "reason": "…" }]`, and must tell the user what was accepted.
4. When it is a G5 `undeclared` difference that the user explicitly wants written anyway, the agent may retry with `override: [{ "id": "d_…", "reason": "…" }]`. The result is `saved_with_overrides`, and the agent must say so. Never override without the user's consent.
5. Failures of G1, G2, G3, G4, G6 and G7 cannot be accepted or overridden. Report them.

## 9. Worked scenarios (legacy defects → new behavior)

| Scenario | Legacy result | New result |
|---|---|---|
| **V-06**: sheet XML uses prefix `x:`; edit `A1` | Edit silently lost; save + verify report success | The DOM writer handles prefixes. If anything goes wrong, G4 fails: `S!A1.value expected "sửa", actual "a"` → `SAVE_BLOCKED` |
| **V-04**: insert a row, and a rich cell further down loses its runs | Hidden by the `worksheets/S/*` exemption | G5 compares `W[T(A39)]` with `O[A39]` → `undeclared` on `S!A40.rich` → blocked |
| **V-02 / V-07**: merge on a sheet with 50 rich cells | Full rebuild → 50 false positives → merge cannot be saved with verification on | No rebuild. Even for files rewritten elsewhere, N-BOOL/N-NUM/N-ORDER classify them as `normalized` |
| **V-03**: merge cells, verify on | `merged_cells` not mapped → always blocked | `merge` declares `S!merge`; G4 checks the new merge list |
| **EX-01**: `.xlsm` edit drops charts, printer settings, the VBA relationship | Success when verify is off | G1 (missing relationship), G6 (VBA content type, charts) fail → blocked |
| **EX-03**: `x14ac:dyDescent` loses its namespace | 20 changes, only with verify on | G3 fails on namespace URI loss |
| **EX-04**: undeclared prefix in `mc:Ignorable` | `valid=true`, `equivalent=true` | G3 fails |
| **EX-05**: row-height edit changes the workbook password hash | Only caught with verify on | G5 `undeclared` on `wb.protection.workbookPassword` |
| **EX-07**: `_xlnm.Print_Titles` loses `hidden="1"` | Not preserved | G5 `undeclared` on `name[_xlnm.Print_Titles@S].hidden` |
| **EX-09**: batch fails midway, earlier cells already changed | Partial application (current code stages per call; re-test) | Atomic by construction (I5) |
| **V-05**: style the first 3 "characters" of NFD "Việt" | Splits a grapheme | Offsets count graphemes; `match` is NFC-insensitive |

### 9.1 End-to-end example: recolor one word

Request:

```json
{ "session": "xs_7k2…", "base_revision": 2, "sheet": "S",
  "ops": [ { "op": "rich_style", "target": "A3", "expect": { "hash": "9f2c41d0be7a" },
             "at": { "match": "đậm" }, "style": { "color": "FF0000FF" } } ] }
```

Response (abridged):

```json
{ "ok": true, "session": "xs_7k2…", "revision": 3,
  "data": {
    "results": [ { "index": 0, "op": "rich_style", "status": "applied", "resolved": "S!A3" } ],
    "diff": [ { "id": "d_41c09e7a2b", "path": "S!A3.rich",
                "before": "<r b color=\"FF0000\">Đỏ đậm </r><r i sz=\"14\">nghiêng </r>…",
                "after":  "<r b color=\"FF0000\">Đỏ </r><r b color=\"0000FF\">đậm</r><r b color=\"FF0000\"> </r><r i sz=\"14\">nghiêng </r>…" } ],
    "readback": { "S!A3": { "value": "Đỏ đậm nghiêng gạch chân 3", "rich": "…same as diff.after…", "hash": "c07d…" } } },
  "warnings": [] }
```

Save report (abridged):

```json
{ "ok": true, "data": { "status": "verified",
  "gates": [ { "id": "G1", "status": "verified" }, { "id": "G2", "status": "verified", "checked": 1 },
             { "id": "G3", "status": "verified" }, { "id": "G4", "status": "verified", "checked": 1 },
             { "id": "G5", "status": "verified", "checked": 4211 }, { "id": "G6", "status": "verified" },
             { "id": "G7", "status": "verified", "checked": 0 } ],
  "readback": { "S!A3": { "rich": "<r b color=\"FF0000\">Đỏ </r><r b color=\"0000FF\">đậm</r>…" } } } }
```

## 10. What "verified" does not mean

- **Formula results:** they are not proven correct unless `recalc` ran; otherwise `cache: "stale"` is reported.
- **Pixels:** pixel-identical display is not proven unless `render` ran, and even then only by the chosen provider.
- **Business logic:** the server proves the file contains what was requested, not that the request was right.
- **Pre-existing problems:** problems already in O (baseline schema errors) are reported, not fixed.
