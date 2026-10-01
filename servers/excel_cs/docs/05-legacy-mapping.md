# 05 — Mapping from the legacy Python server

The legacy server is `servers/excel` (Python, FastMCP, openpyxl; 112 tools registered at runtime, `server.json` 1.1.2). This document lists **every** legacy tool and mechanism, says how it works today, and says what the C# server provides instead and what must change.

Line references point at the legacy code as of this design (`main.py` 9,010 lines, `core.py` 6,597, `preservation.py` 2,016).

## 1. Summary

| | Legacy | New |
|---|---|---|
| Tools | 112 | 15 |
| Mutation entry points | ~70 tools, each staging its own change | 1 tool (`excel_apply`) with ~75 typed ops, transactional per call |
| Addressing | 0-based `row_index`/`col_index`/`r1,c1`, A1 in some tools, exclusive `end_row`, 1-based `shape_index` | A1 only, 1-based rows, inclusive ranges, stable object ids |
| Output | Mostly `json.dumps` inside text; `excel_load` returns text with `session_key='…'`; save report is text by default | `structuredContent` with an `outputSchema`, plus a compact JSON text block |
| Save | 3 tiers: exact copy / regex patch of edited cells / openpyxl rebuild + 17 `_inject_*` passes | 1 path: PackageStore writer, DOM edits only on touched parts |
| Verification | Optional (`verify_preservation=False`), two-way, glob-based exemptions | Mandatory, three-way, exact effects + transforms |

## 2. Conventions

| Topic | Legacy | New | Conversion |
|---|---|---|---|
| Row index | 0-based (`row_index=4` = Excel row 5) | Excel row number | `new = old + 1` |
| Column index | 0-based (`col_index=2` = C) | Letter | `new = letter(old + 1)` |
| Ranges in delete_rows | `start_row=14, end_row=19` (end exclusive, 0-based) = Excel rows 15–19 | `"15:19"` | `"{start+1}:{end}"` |
| Rectangles | `r1,c1,r2,c2` (0-based, inclusive) | `"B2:D10"` | — |
| Markdown export | `col_0…`, 0-based `row_index` annotations | A, B, C… headers, Excel row numbers | — |
| Shapes | `shape_index` 1-based, positional | `dr_…` id from `excel_inspect aspect=drawings` | — |
| CF rules | per-session `rule_id` | `cf_…` content-derived id, stable across revisions | — |
| Session key | file path, or `new:<token>:<ext>`; parsed from text | `xs_…` in structured output | — |
| Formulas | Stored/returned with `=` in the public model; without `=` in raw XML | Always returned with `=`; accepted with or without | — |
| Rich text | runs with code-point `start`/`end` | Markup + `match`/grapheme spans | — |
| Multi-insert | Clone all template rows first, then insert bottom-to-top | `address_mode: "base"`; the server handles drift | — |
| Errors | Mostly `ValueError` text; a few structured `EXCEL_*` codes | Structured `error.code` everywhere (03 §7) | §5 |

## 3. Tool mapping

Columns:
- **Legacy behavior:** what the tool does today, including relevant defects.
- **New:** the tool, or `excel_apply` op, that replaces it.
- **New requirement:** what must be different.

### A. Session and file (7)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_load(uri, sheet_name?, timeout_seconds?)` | Cancellable worker serializes the **whole workbook** to a dict and ships it as JSON. `sheet_name` loads one sheet, and then formulas on unloaded sheets are not rewritten by structural edits. Returns text containing `session_key='…'` | `excel_open` | Lazy part loading, so no sheet filter is needed; structured overview; feature warnings (signatures, VBA, NFD text); `resume` |
| `excel_create_workbook(format, sheet_names, active_sheet, template_path, macro_template_path, document_properties, target_path)` | New in-memory dict model; key `new:<token>:<ext>` | `excel_create` | `vba_from` replaces `macro_template_path`; first save runs G1–G4 + G7 |
| `excel_close(session_key)` | Drops the session even when it has unsaved edits | `excel_close` | `UNSAVED_CHANGES` unless `discard_unsaved: true` |
| `excel_get_session_status(session_key)` | Busy flag, dirty feature/path counts | `excel_status` | Adds ledger, `source_changed_on_disk`, oracle availability, list mode |
| `excel_reload(session_key)` | Re-reads the file from disk, discarding edits. Untested per audit | `excel_undo {to_revision: saved_revision}`; if the disk changed: `excel_close` + `excel_open` | Never silently loads a file that changed on disk |
| `excel_save(session_key, output_path?, report_format="text", verify_preservation=False, max_differences=200, timeout_seconds?)` ([main.py:3139](../../excel/main.py#L3139)) | Picks one of three tiers ([core.py:5852](../../excel/core.py#L5852)); verification **off by default**, except for row-only edits; text report by default | `excel_save {mode: "overwrite" \| "save_as"}` | Gates always on; structured report; readback from the written file; `assert`, `accept`, oracles |
| `excel_save_as_copy(session_key, output_path, …)` | Same pipeline to another path; the path must differ from the source; the session's dirty state is cleared afterwards | `excel_save {mode: "copy"}` | The session stays on its source and **stays dirty**. Use `save_as` to move the session |

### B. Stateless (4)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `convert_to_markdown(file_path, sheet_name?, range_ref?, max_rows?, max_cols?, include_styles=False)` | openpyxl read; error hints suggest it for `.xls`/`.xlsb`, which does not work (audit) | `convert_to_markdown` (same signature) | A1 headers; honest format errors; optional ExcelDataReader for legacy formats |
| `excel_get_info(uri)` | Sheet names + row/col counts | `excel_peek {detail: "info"}` | — |
| `excel_get_sheet_preview(file_path, max_rows=20, max_cols=10, sheet_name?)` | Top-left previews | `excel_peek {detail: "preview"}` | A1 headers |
| `excel_get_workbook_summary(file_path)` | Compact summary | `excel_peek {detail: "summary"}` | Includes feature flags as in `excel_open` |

### C. Reading (9)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_get_cell(sheet, row_index, col_index, include_rich_text, include_formula_cache, include_semantics)` | One 0-based cell; optional semantic fields | `excel_read {target: "B5", view: "cells" \| "full"}` | Display string, effective style, markup, hash |
| `excel_get_rows(sheet, start_row=0, end_row?, values_only, include_*)` | 0-based row slice | `excel_read {target: "5:20"}` | Pagination by `max_cells`/`cursor` |
| `excel_get_column(sheet, col_index, start_row, end_row?)` | Column as JSON | `excel_read {target: "C2:C100"}` | — |
| `excel_read_range(sheet, range_ref \| start/end row/col, values_only=True, include_*)` | Rectangle, A1 or 0-based | `excel_read` | One addressing form |
| `excel_get_rich_text(sheet, cell)` | Runs with code-point offsets, whitespace, phonetics | `excel_read {target, include: ["rich"]}` | Markup; grapheme semantics; `text_form` in `full` view |
| `excel_to_markdown(session_key, sheet?, max_rows?, max_cols?)` | Markdown annotated with 0-based indices. Untested per audit | `excel_read {view: "markdown"}` | A1 headers, `†`/`ƒ` markers |
| `excel_to_markdown_range(sheet, range_ref \| bounds)` | Range Markdown | `excel_read {view: "markdown", target}` | — |
| `excel_find_cells(query, sheet?, regex, case_sensitive, match_in="value", max_results=100)` | Literal/regex search | `excel_find` | Search in `display`/`value`/`formula`/`rich`/`comment`; NFC-insensitive; style predicates; grapheme spans |
| `excel_find_rows(sheet, col_index, value \| pattern, start_row, end_row?)` | Rows where a column matches. Untested per audit | `excel_find {scope: {target: "C:C"}}` | Rows are derived from `addr` |

### D. Inspection and export (11)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_get_conditional_formats(sheet, sqref?, include_raw_xml)` | Rules with per-session ids | `excel_inspect {aspect: "conditional_formats"}` | Content-derived `cf_` ids; effective dxf style |
| `excel_get_shapes(sheet?)` | DrawingML shapes/images/charts captured at load | `excel_inspect {aspect: "drawings"}` | `dr_` ids; A1 anchors; shape text as markup; chart series refs |
| `excel_get_sheet_semantics(sheet)` | State, properties, views, printing, breaks, protection, errors | `excel_inspect {aspect: "sheet"}` | — |
| `excel_get_sheet_views(sheet)` | Sheet views | `excel_inspect {aspect: "sheet"}` | — |
| `excel_get_workbook_semantics()` | Calc/properties/protection/views/doc props | `excel_inspect {aspect: "workbook"}` | Protection reported as flags and algorithm, never as secrets |
| `excel_get_workbook_views()` | Workbook views | `excel_inspect {aspect: "workbook"}` | — |
| `excel_list_defined_names()` | Names from the session | `excel_inspect {aspect: "names"}` | — |
| `excel_list_package_parts(prefix?, max_parts=200, max_relationships_per_part=20, include_deleted)` | Effective parts with hashes and relationship edges | `excel_inspect {aspect: "package"}` | — |
| `excel_list_tables(sheet?)` | Tables in the session | `excel_inspect {aspect: "tables"}` | — |
| `excel_read_package_part(part_path, output_mode="auto", offset, max_bytes=65536)` | Bounded XML/text/bytes/base64 | `excel_inspect {aspect: "part"}` | — |
| `excel_extract_images(sheet, output_dir)` | Writes embedded images to disk. Untested per audit | `excel_export {what: "images"}` | Returns sha256 per file |

### E. Cell content (5)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_edit_cells(sheet, edits[])` ([main.py:5357](../../excel/main.py#L5357)) | Three input forms (A1, flat 0-based, grouped by row); typed payloads; on a rich cell, plain values require `rich_text_policy` (`replace_all` / `preserve_runs_if_text_equal`). Returns full before/after cell views per cell. Audit EX-09 found partial application; current code stages per call, so this needs re-testing | `set_value`, `set_values`, `set_formula`, `rich_set` | One A1 form; `rich_policy: "reject" \| "replace"`; `AMBIGUOUS_FORMULA_TEXT`; sparse diff; atomic across the whole `excel_apply` |
| `excel_edit_rich_text(sheet, cell, operations[], expected_text?)` | Ops `replace_runs`, `style_run(run_index)`, `style_range`/`insert_text`/`delete_range`/`replace_range` (code-point `start`/`end`), `set_phonetic`. Always rewrites the cell as `inlineStr` | `rich_set`, `rich_style`, `rich_replace`, `rich_insert`, `rich_delete`, `phonetic_set` | `match` or grapheme spans (V-05); `expected_text` becomes `expect.text`/`expect.hash`; storage form kept |
| `excel_set_formula(sheet, cell, formula, formula_type="normal", formula_attributes?, cached_value?, cached_value_present, cache_policy="clear")` | Normal/shared/array/dataTable with explicit cache state | `set_formula` | Declares `fullCalcOnLoad` and calcChain effects; detaches shared members explicitly |
| `excel_clear_range(sheet, r1, c1, r2, c2, clear_values=True, clear_styles=False)` | 0-based rectangle | `clear {target, what}` | Also comments/hyperlinks/validation; `remove_cells` |
| `excel_fill_column(sheet, col_index, start_row, end_row, value \| sequence_start, step=1)` | Constant or sequence | `fill {target: "C2:C20", value \| series}` | Works on any range, not only one column |

### F. Style (8)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_set_style(sheet, r1, c1, r2?, c2?, style={})` | Patches whole-cell font/fill/alignment/protection/XF; rich runs stay intact | `style {target, patch}` | Sparse patch semantics (02 §3.3); `apply_to_runs` option |
| `excel_set_font_color(sheet, r1, c1, color, r2?, c2?)` | ARGB or null | `style {patch: {font: {color}}}` | Theme/indexed colors accepted |
| `excel_set_strike(sheet, r1, c1, enabled=True, r2?, c2?)` | — | `style {patch: {font: {strike}}}` | — |
| `excel_set_borders(sheet, r1, c1, r2, c2, style?, sides?, color?, border?)` | Patches sides, keeps unspecified siblings | `border {outline, inside, sides}` | Range outline vs inner edges |
| `excel_set_cell_style_semantics(sheet, range_ref, xf?, named_style?, exact_default_policy="preserve", dry_run)` | Expert XF flags/base style | `style_xf` (expert) | `dry_run` is general in `excel_apply` |
| `excel_add_named_style(name, style, metadata?)` | No renumbering of existing styles | `named_style_add` | — |
| `excel_update_named_style(name, updates)` | Keeps identity | `named_style_update` | — |
| `excel_delete_named_style(name)` | Only when unused | `named_style_delete` | `IN_USE` with references |

### G. Rows, columns and merges (14)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_insert_rows(sheet, inserts[{after_index, rows_json}])` | Rows come as JSON from `excel_clone_rows`; positions sorted bottom-to-top. Shifts merges/links/CF/DV/tables/anchors/formulas on loaded sheets. Verification exempts the whole sheet (V-04) | `insert_rows`, or `copy_rows {source}` | Transform-aware G5; predicted rewrites listed; independent shifter (04 §6.2) |
| `excel_delete_rows(sheet, row_indices? \| start_row+end_row)` | 0-based; `end_row` exclusive | `delete_rows {rows: "15:19"}` | Inclusive A1; `#REF!` rewrites listed |
| `excel_clone_rows(sheet, start_row, end_row?)` | Returns rows as JSON **without** inserting; the agent passes them back | — (not needed) | `copy_rows` copies server-side; no row JSON round-trip |
| `excel_copy_row(sheet, row_index, after_index)` | Clone + insert in one step | `copy_rows {source: "N", after}` | — |
| `excel_fill_rows(sheet, template_row, after_index, count)` | Clone a template N times | `copy_rows {source, after, times}` | — |
| `excel_insert_column(sheet, after_col_index)` | One empty column | `insert_cols {after: "C", count}` | Count supported |
| `excel_delete_column(sheet, col_index)` | One column | `delete_cols {cols: "C:D"}` | Ranges supported |
| `excel_copy_column(sheet, col_index, after_col_index)` | Copy with styles | `copy_cols` | — |
| `excel_set_row_height(sheet, row_heights{idx: h})` | Dedicated patch path `_stage_row_heights`; the verifier exempts the whole row (audit) | `row_props {rows, height}` | Only `S!row:N.height` is declared |
| `excel_set_row_properties(sheet, row_index, properties)` | Height/hidden/outline/collapse/style/phonetic | `row_props` | Only the given attributes are declared |
| `excel_set_column_width(sheet, col_widths)` | Several columns | `col_props {cols, width}` | — |
| `excel_set_dimension(sheet, axis, index, size)` | Row height or column width | `row_props` / `col_props` | — |
| `excel_autofit_cols(sheet, col_indices?, min_width=8, max_width=60)` | Heuristic based on content length | `col_props {width: "autofit", autofit: {min, max}}` | Still heuristic (font metrics); widths declared as effects; `excel_render` to confirm |
| `excel_merge_cells(sheet, r1, c1, r2?, c2?, unmerge=False)` | Merge/unmerge; `sheets/X/merges` is not mapped to verifier paths, so merging with verification on is always blocked (V-03) | `merge` / `unmerge` | `on_data_loss` policy; declares `S!merge` |

### H. Sheets (9)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_add_sheet(sheet_name, position?)` | Blocked by the verifier: new relationships/content types are not requested (EX-10) | `sheet_add` | New parts, relationships and content types are declared effects |
| `excel_delete_sheet(sheet_name)` | Not the only sheet | `sheet_delete` | `on_references` policy; `#REF!` listed |
| `excel_rename_sheet(sheet_name, new_name)` | — | `sheet_rename` | Reference rewrites declared |
| `excel_move_sheet(sheet_name, position)` | — | `sheet_move` | `localSheetId` and active tab updated and declared |
| `excel_copy_sheet(source_sheet, new_name, position?)` | Within the workbook | `sheet_copy` | Related parts copied and declared |
| `excel_copy_sheet_to(src_session_key, src_sheet_name, dst_session_key, new_name?, position?)` | Across sessions | `sheet_copy_from` (in the destination session) | Style remapping declared |
| `excel_set_sheet_state(sheet_name, state)` | Keeps one sheet visible | `sheet_state` | — |
| `excel_set_sheet_properties(sheet_name, properties)` | CodeName/filter/published/sync/outline/pageSetup props | `sheet_props` | — |
| `excel_set_ignored_errors(sheet_name, rules, mode="patch")` | — | `ignored_errors` | — |

### I. Data features (15)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_add_table(sheet, name, ref, table?, style?)` | Blocked by the verifier (EX-10) | `table_add` | Table part and relationships declared |
| `excel_update_table(name, updates)` | Reconciles range-bound metadata | `table_update` | — |
| `excel_delete_table(name)` | — | `table_delete {keep_data}` | — |
| `excel_set_auto_filter(sheet, ref?, filter_columns?, sort_state?, mode="patch")` | — | `autofilter` | `_FilterDatabase` name declared |
| `excel_add_conditional_format(sheet, sqref, rule)` | cellIs/expression/colorScale/dataBar/iconSet | `cf_add` | — |
| `excel_update_conditional_format(sheet, rule_id, updates)` | Keeps omitted XML siblings | `cf_update` | — |
| `excel_delete_conditional_format(sheet, rule_id)` | Compacts remaining priorities | `cf_delete` | Priority changes declared |
| `excel_set_data_validation(sheet, start/end row/col?, options?, allow_blank, validation?, sqref?, mode="append")` | Add/replace/patch through one tool | `dv_add` / `dv_update` / `dv_delete` | `dv_` ids |
| `excel_add_defined_name(name, value, sheet_name?, local_sheet_id?, metadata?)` | Rejects `_xlnm.*` built-ins | `name_add` | Same built-in rule (print/filter ops own them) |
| `excel_update_defined_name(name, updates, …)` | Keeps omitted metadata | `name_update` | — |
| `excel_delete_defined_name(name, …)` | Exact scope | `name_delete` | — |
| `excel_set_hyperlink(sheet, cell, target?, location?, display?, tooltip?, relationship?)` | Blocked by the verifier for new targets (EX-10) | `hyperlink_set` | Relationship declared |
| `excel_remove_hyperlink(sheet, cell)` | — | `hyperlink_remove` | Orphaned relationship removal declared |
| `excel_set_comment(sheet, cell, text, author="", comment_type="legacy", metadata?)` | Legacy comments only | `comment_set` | Threaded comments are read-only; preserved by G6 |
| `excel_remove_comment(sheet, cell)` | — | `comment_remove` | — |

### J. Printing (6)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_set_print_area(sheet, areas)` | Multi-area through the defined-name model | `print_area` | — |
| `excel_set_print_titles(sheet, repeated_rows?, repeated_columns?)` | Built-in name; loses `hidden="1"` (EX-07) | `print_titles` | Keeps all name attributes |
| `excel_set_page_setup(sheet, properties, present?, exact=False)` | Keeps explicit-empty presence | `page_setup` | — |
| `excel_set_print_options(sheet, properties, present?, exact=False)` | — | `print_options` | — |
| `excel_set_header_footer(sheet, sections, properties?)` | Keeps sibling sections | `header_footer` | — |
| `excel_set_page_breaks(sheet, row_breaks?, column_breaks?, mode="replace", exact_attributes?)` | Syncs counts | `page_breaks` | Counts covered by N-COUNT |

### K. Views and protection (6)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_freeze_panes(sheet, row, col)` | Freezes above `row` / left of `col` | `freeze {at: "B2"}` | A1 semantics: rows above 2, columns left of B |
| `excel_set_sheet_views(sheet, views, mode="replace")` | — | `sheet_views` | — |
| `excel_set_workbook_views(views, mode="replace")` | — | `workbook_views` | — |
| `excel_set_sheet_protection(sheet, properties?, enabled?, already_hashed=True)` | — | `sheet_protection` | `password` (server hashes) or `hash` |
| `excel_set_protected_ranges(sheet, ranges, mode="replace")` | — | `protected_ranges` | — |
| `excel_set_workbook_protection(properties, already_hashed=True)` | Unrelated edits changed the password hash (EX-05) | `workbook_protection` | G5 catches any undeclared protection change |

### L. Workbook properties (3)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_set_workbook_properties(properties, date_system_policy?)` | `date1904` change needs a policy | `workbook_props` | Same rule |
| `excel_set_calculation_properties(properties)` | Keeps explicit false/zero | `calc_props` | — |
| `excel_set_document_properties(core?, app?, custom?, modified_policy="preserve")` | Typed custom properties | `doc_props` | — |

### M. Drawings (5)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_add_image(sheet, anchor, source_path? \| base64_data?, mime_type?, width?, height?, name?)` | Queued, created at save | `image_add` | Created in the overlay at apply time, so readback works |
| `excel_add_chart(sheet, chart_type, source_range, anchor, title?, width?, height?, options?)` | Queued | `chart_add` | Same |
| `excel_add_shape(sheet, shape_type, anchor, text?, rich_text?, width?, height?, style?, name?)` | Queued | `shape_add` | Shape text as markup |
| `excel_update_shape_text(sheet, shape_index, text?, rich_text?)` | 1-based index; regex over the drawing XML | `shape_text {id}` | DOM edit |
| `excel_set_shape_style(sheet, shape_index, fill_color?, outline_color?, outline_width_pt?, text_color?, clear_*)` | 1-based index | `shape_style {id}` | — |
| — | No way to delete a drawing | `drawing_delete` | New |

### N. Package (5)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_apply_package_transaction(upsert?, delete?, relationships?, content_types?)` | Atomic staging of package edits | `part_upsert` / `part_delete` / `rels_set` / `content_types_set` in one `excel_apply` (expert) | Same atomicity, via `excel_apply` |
| `excel_upsert_package_part(part_path, content, encoding="text", content_type?)` | Applied at save | `part_upsert` | Applied in the overlay; G2 validates it if it has a schema |
| `excel_delete_package_part(part_path)` | Dangling-reference validation | `part_delete` | `IN_USE` with referencing parts |
| `excel_set_package_relationships(source_part, relationships)` | Replaces all relationships of a part | `rels_set` | — |
| `excel_set_package_content_types(defaults?, overrides?)` | — | `content_types_set` | — |

### O. Verification, evidence and rendering (5)

| Legacy tool | Legacy behavior | New | New requirement |
|---|---|---|---|
| `excel_validate_workbook(path)` | ZIP + XML well-formedness + feature report; `valid=true` for an undeclared `mc:Ignorable` prefix (EX-04) | `excel_verify {after_path}` | G1–G3 including schema and MC checks |
| `excel_diff_package(before_path, after_path)` | ZIP manifest diff | `excel_verify {…, detail: "package"}` | — |
| `excel_verify_preservation(after_path, before_path?, max_differences, requested_paths?, approved_normalizations?, fixture_gap_paths?, verifier_gap_paths?, fixture_id?)` | Two-way semantic diff in a worker; glob `requested_paths`; caller-supplied normalizations | `excel_verify` | `session` for intent; `accept` by id instead of patterns; fixture concepts move to the evidence CLI |
| `excel_build_preservation_summary(coverage_reports, verification_reports, backup_checks, …)` | Aggregates evidence JSON | `tools/evidence` CLI | Not an agent tool |
| `excel_capture(sheet, output_path, soffice_path?, timeout_seconds=120)` | Renders a sheet to PNG via LibreOffice; kills the process tree on cancel | `excel_render` | Ranges, providers (Excel COM preferred), image content returned to the agent |

**Count check:** A 7 + B 4 + C 9 + D 11 + E 5 + F 8 + G 14 + H 9 + I 15 + J 6 + K 6 + L 3 + M 5 + N 5 + O 5 = **112**.

## 4. Mechanism mapping

| Legacy mechanism | Where | New |
|---|---|---|
| Workbook model as a dict (`serialize_excel`); worker artifact; `deepcopy` per edit; `_checkpoint_session` | `core.py`, `main.py`, `cancellable.py` | PackageStore overlay + ledger (01 §4) |
| Dirty tracking: `_dirty_features`, `_dirty_paths` via `_mark_dirty` | [main.py:1014](../../excel/main.py#L1014) | Ledger effects, exact semantic paths |
| Verifier exemptions: `_save_verifier_patterns` turns dirty paths into globs (`worksheets/S/cells/B5*`, `worksheets/S/*` for structure) | [main.py:2667](../../excel/main.py#L2667), rows branch [main.py:2746](../../excel/main.py#L2746) | Declared effects + transforms; no globs |
| Save tier 1: `_atomic_copy_source_package` (no edits) | `core.py` | Same idea: no ledger entries → copy bytes, then G1 |
| Save tier 2: `_content_only_changes` + `_reconstruct_content_only` (regex `_find_cell_xml` / `_insert_cell_xml`) | [core.py:4232](../../excel/core.py#L4232), [core.py:4017](../../excel/core.py#L4017) | Removed: DOM edit of touched parts (fixes V-06) |
| Save tier 3: openpyxl rebuild + 17 `_inject_*` + `_restore_missing_package_parts` | [core.py:5852](../../excel/core.py#L5852) | Removed: no rebuild path exists (fixes V-02, EX-01, EX-03) |
| Row-height fast path `_stage_row_heights` | `save_transaction.py` | Not needed: normal DOM edit |
| Verifier `verify_xlsx_preservation` (ElementTree snapshot, two-way, `_Changes` classifier) | `preservation.py` ([_compare_cells](../../excel/preservation.py#L1581), [_rich_text](../../excel/preservation.py#L1093)) | `Verify` module: independent reader, three-way, normalization registry |
| Signature check `package_signature_report`; `EXCEL_SAVE_REQUIRES_RESIGNING` | `preservation.py`, `save_transaction.py` | G6 + `allow_signature_invalidation` |
| Advanced part guard `_guard_advanced_package_parts` | `save_transaction.py` | G6 (all advanced parts, semantically) |
| Staging, `file_state`/`_assert_state`, `commit_staging_file`, backups with 2-day retention | `save_transaction.py`, `preservation.py` | Kept (01 §5.3) |
| Worker processes for load/save/verify; `EXCEL_MCP_*` limits | `cancellable.py` | In-process cancellation; same limit names |
| Formula/reference shifting on structural edits | `core.py` (tests: `test_excel_formula_shift.py`) | `Formula` module + independent shifter in `Verify` |
| Implicit cells, lazy baselines, cell baseline hashes | `main.py`, `core.py` (tests: `test_excel_implicit_cells.py`, `test_excel_lazy_baseline.py`) | Not needed: no materialized model |
| Package tools registered via `register_package_tools` | `package_tools.py` | Expert ops in `excel_apply` |
| Server instructions with index conventions and the multi-insert pattern | `main.py` | Replaced (03 §1.3) |

## 5. Error codes

| Legacy code | New code | Note |
|---|---|---|
| `EXCEL_SESSION_BUSY` | `SESSION_BUSY` | — |
| `EXCEL_LOAD_RESULT_INVALID` | — | No worker envelope any more |
| `EXCEL_SAVE_PRESERVATION_FAILED` | `SAVE_BLOCKED` | Details carry per-gate results |
| `EXCEL_SAVE_REQUIRES_RESIGNING` | `SIGNATURE_WOULD_BREAK` | — |
| `EXCEL_SAVE_ADVANCED_PARTS_LOST` | `SAVE_BLOCKED` (G6) | — |
| `EXCEL_SAVE_CONCURRENT_FILE_CHANGE` | `SOURCE_CHANGED_ON_DISK` / `DESTINATION_CHANGED` | Split by which file changed |
| `EXCEL_SAVE_BACKUP_HASH_MISMATCH` | `BACKUP_FAILED` | — |
| `EXCEL_SAVE_STAGING_MISSING`, `EXCEL_SAVE_STAGING_CLEANUP_FAILED` | `INTERNAL_ERROR` (with `details.stage`) | — |
| `EXCEL_SAVE_SESSION_REPLACED` | `REVISION_CONFLICT` | — |
| `EXCEL_SAVE_SHEET_NOT_FOUND`, `EXCEL_SAVE_ROW_NOT_FOUND`, `EXCEL_SAVE_DUPLICATE_ROW`, `EXCEL_SAVE_SHEET_DATA_NOT_FOUND` | — | Internal invariants of the regex/row patch paths, which no longer exist |
| Unstructured `ValueError` messages (most input errors) | `INVALID_INPUT`, `INVALID_OP`, `INVALID_ADDRESS`, `TARGET_*`, `PRECONDITION_FAILED`, … | 03 §7 |

## 6. Environment variables

See [01 §10](01-architecture.md#10-configuration-summary). These variables keep their names:

- `EXCEL_MCP_LOAD_TIMEOUT_SECONDS`, `EXCEL_MCP_SAVE_TIMEOUT_SECONDS`, `EXCEL_MCP_VERIFY_TIMEOUT_SECONDS`, `EXCEL_MCP_HEAVY_TIMEOUT_SECONDS`
- `EXCEL_MCP_MAX_HEAVY_WORKERS`, `EXCEL_MCP_MAX_INPUT_BYTES`, `EXCEL_MCP_MAX_RESULT_BYTES`
- `DOCLOUPE_EXCEL_BACKUP_DIR`

`EXCEL_MCP_MAX_LOAD_RESULT_BYTES` and `EXCEL_MCP_WORKER_PROCESS` are dropped.

## Known defect register

**V-xx** were found while preparing this design: code reading plus probe scripts against the legacy server. **EX-xx** come from the audit `EXCEL_TOOLS_AUDIT_2026-09-30.md`.

| ID | Severity | Legacy defect | Evidence | Addressed by |
|---|---|---|---|---|
| V-01 | High | `verify_preservation=False` by default on `excel_save`/`excel_save_as_copy`; only a warning is emitted | [main.py:3139](../../excel/main.py#L3139); probe: `status: not_run` | D5: gates always on |
| V-02 | Medium | Rich-run properties compared as raw XML: `<b val="1"/>` vs `<b/>`, `sz="14"` vs `"14.0"`, child order. Merging one range on a sheet with 50 rich cells gave 50 false `UNAPPROVED_LOSS` results | [preservation.py:1093](../../excel/preservation.py#L1093); probe after merge | No rebuild path; N-BOOL, N-NUM, N-ORDER |
| V-03 | Medium | `sheets/X/merges` has no branch in `_save_verifier_patterns`, so a merge with verification on is always blocked | [main.py:2667](../../excel/main.py#L2667); probe | `merge` declares `S!merge` |
| V-04 | **Critical** | Row/column insert/delete exempts `worksheets/S/*`. 204 changes were all `REQUESTED`; a real loss on that sheet would pass | [main.py:2746](../../excel/main.py#L2746); probe | Transform-aware G5 (04 §6) |
| V-05 | Medium | Rich-text offsets in code points; `style_range` on NFD "Việt" splits a grapheme (`"Vie"` / `"̣̂t"`) | `_chars_from_rich` in `main.py`; probe | Grapheme offsets, NFC-insensitive `match` |
| V-06 | **Critical** | Sheet XML with a namespace prefix (`<x:c>`): the content-only regex patch does not find the cell, the edit is dropped, and save + verify report success with 0 changes | [core.py:4017](../../excel/core.py#L4017) and `_find_cell_xml`; probe | DOM writer; G4 intent check |
| V-07 | Low | Row attributes compared as strings (`ht="30"` vs `"30.0"`), so the verifier blocks a valid save | Probe (self-closing row case) | N-NUM |
| V-08 | **Critical** | No check that declared edits are present in the written file | `verify_xlsx_preservation` design | G4 |
| V-09 | Medium | JSON inside text; `session_key` parsed from text; full before/after cell views per edited cell (token-heavy) | `excel_load`, `_mutation_result`, `excel_edit_cells` | `structuredContent`; sparse diffs |
| V-10 | Medium | Mixed addressing (0-based `r1/c1`, A1, exclusive `end_row`, 1-based `shape_index`) | Tool signatures | A1 only |
| V-11 | Low–medium | One cell view contains keys that differ only by case (`vAlign` = font vertical alignment, `valign` = cell vertical alignment). Windows PowerShell 5.1 `ConvertFrom-Json` refuses the payload; other case-insensitive clients may merge the keys | Clean-machine harness run, 2026-10-01 (`tests/clean-machine/windows-sandbox`) | Key rule in 03 §1.1 |
| V-12 | Info | Startup: about 18–20 s from process start to the `initialize` response (PyInstaller one-file extraction + imports), measured on the development machine for both the local `dist` build and the released v1.1.2 binary | Same harness run | Native AOT target < 150 ms (01 §12.2) |
| EX-01 | Critical | `.xlsm` row-height edit: 36 → 31 parts; charts, PNG, printer settings, calcChain lost; VBA relationship and macroEnabled content type lost | Audit | G1, G6; no rebuild |
| EX-02 | High | `07-real-package-source.xlsx`: save fails with `unbound prefix` | Audit | DOM serialization keeps declarations; G3 |
| EX-03 | High | `x14ac:dyDescent` loses its namespace; 20 changes outside the request | Audit | G3; no attribute-name stripping |
| EX-04 | High | Undeclared prefix in `mc:Ignorable`: `valid=true`, `equivalent=true` | Audit | G3 |
| EX-05 | High | Row-height edit changes the workbook password hash, views, names, styles, properties, filter | Audit | G5 |
| EX-06 | Medium–high | Rich-text font/run differences in D1/D3, `xml:space` added | Audit | No rebuild; N-XMLSPACE with a strict condition |
| EX-07 | Medium | `_xlnm.Print_Titles` loses `hidden="1"` | Audit | G5 on `name[…].hidden` |
| EX-08 | Medium | Table filter gains `showButton`/`hiddenButton` defaults; table relationship id changes | Audit | N-DEFAULT, N-RID |
| EX-09 | Medium | Batch error leaves earlier cells changed | Audit (re-test against current code) | Atomic `excel_apply` |
| EX-10 | Medium | Adding a sheet, hyperlink or table is blocked: new relationships/content types are not requested | Audit | Package effects declared by the ops |
