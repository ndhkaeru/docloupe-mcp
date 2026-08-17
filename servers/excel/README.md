# excel-tools

Round-trip Excel workbook editor with one-shot Markdown export plus session-based reads, sheet operations, rows/columns, styles, merges, data validation, best-effort DrawingML shape/image/chart preservation, and validated save/reload support.

Use this server when an agent needs Excel content without dumping an entire workbook into context. The tools can inspect only the needed sheets, rows, columns, or cells, which helps reduce token usage and quota consumption.

## Status

Implemented and included in the completed DocLoupe MCP server set. The preservation regression suite currently passes `175` tests with no xfails, and all seven immutable preservation fixtures complete no-edit public round trips byte-identically with zero unapproved semantic differences.

## Input

- `convert_to_markdown` is read-only and can convert Excel-family files readable by the converter into Markdown without creating a session.
- Session/edit/save tools support OOXML Excel packages (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) with best-effort macro/template preservation. Legacy/binary formats such as `.xls` and `.xlsb` can still be used through read-only conversion paths when the converter supports them.
- Save is validated before replacing the destination file. If the generated workbook package is invalid, the existing destination is left untouched.
- Advanced Excel features such as DrawingML shapes, charts, images, VML drawings, pivot/slicer parts, external links, and unknown OOXML parts are preserved best-effort. Not every advanced object is editable.
- For non-Excel document types, use the matching DocLoupe MCP server: Markdown, PDF, DOCX, PPTX, CSV, HTML, text, JSON, or JSONL.

## Recommended Workflow

- Start with `excel_get_info`, `excel_get_workbook_summary`, or `excel_get_sheet_preview` for compact workbook context.
- Use `convert_to_markdown` when you only need a Markdown export and do not need a session; pass `sheet_name`, `range_ref`, `max_rows`, or `max_cols` to keep output small.
- Prefer targeted reads such as `excel_read_range`, `excel_find_cells`, `excel_get_rows`, `excel_get_cell`, and `excel_get_column` to avoid spending tokens on irrelevant workbook content.
- Call `excel_load` and keep the returned `session_key`.
- Inspect with `excel_to_markdown`, `excel_get_rows`, or cell/column tools.
- Mutate the session, then call `excel_save` to write an `.xlsx`.

## Tools

| Tool                         | Description                                                                                                                                                                                                          |
| :---------------------------- | :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `convert_to_markdown`        | Convert an Excel-family file to Markdown in one call without creating a session; supports sheet/range/max row-column limits.                                                                                         |
| `excel_get_workbook_summary` | Return a compact read-only workbook summary without creating a session.                                                                                                                                              |
| `excel_get_sheet_preview`    | Return compact top-left previews for one sheet or all sheets without creating a session.                                                                                                                             |
| `excel_get_info`             | Return summary info about an Excel file: sheet names, row and column counts.                                                                                                                                         |
| `excel_load`                 | Load an Excel file through a cancellable worker, validate a bounded tagged-JSON artifact, and publish the session only after complete success.                                                                        |
| `excel_save`                 | Transactionally reconstruct through a cancellable worker, optionally verify staging, then create a two-day backup and atomically replace the destination.                                                            |
| `excel_save_as_copy`         | Use the same cancellable transactional save contract for another OOXML path while leaving the source unchanged.                                                                                                      |
| `excel_validate_workbook`    | Validate an OOXML Excel ZIP/XML package and report advanced features.                                                                                                                                                |
| `excel_diff_package`         | Compare OOXML package manifests and semantic part hashes.                                                                                                                                                            |
| `excel_verify_preservation`  | Run cancellable semantic/package verification in an isolated worker with optional `timeout_seconds`, file health, runtime metadata, six-way difference classification, bounded previews, and latest-backup fallback. |
| `excel_reload`               | Reload session data from disk, discarding any unsaved in-memory changes.                                                                                                                                             |
| `excel_close`                | Remove a session from the server cache to free memory.                                                                                                                                                               |
| `excel_to_markdown`          | Export session data as Markdown tables annotated with 0-based row/column indices; supports `max_rows` and `max_cols`.                                                                                                |
| `excel_to_markdown_range`    | Export one worksheet range as a compact Markdown table.                                                                                                                                                              |
| `excel_list_tables`          | List Excel table objects captured in the loaded workbook session.                                                                                                                                                    |
| `excel_list_defined_names`   | List workbook defined names and named ranges from the loaded session.                                                                                                                                                |
| `excel_capture`              | Render a sheet as a PNG image using LibreOffice.                                                                                                                                                                     |
| `excel_extract_images`       | Extract all embedded images from a sheet and save them to a directory.                                                                                                                                               |
| `excel_get_shapes`           | List captured DrawingML shape/image/chart metadata, including rich-text runs for shape text.                                                                                                                         |
| `excel_update_shape_text`    | Replace captured shape text with a plain string or the common rich-text run model.                                                                                                                                   |
| `excel_set_shape_style`      | Set fill, clear fill, outline color/width, clear outline, and text color for a simple DrawingML shape/textbox.                                                                                                       |
| `excel_get_rows`             | Get a range of rows from session data as JSON.                                                                                                                                                                       |
| `excel_read_range`           | Read an exact rectangular range from a loaded worksheet.                                                                                                                                                             |
| `excel_get_cell`             | Get full metadata of a single cell.                                                                                                                                                                                  |
| `excel_get_column`           | Get all cells in a column as JSON.                                                                                                                                                                                   |
| `excel_find_cells`           | Find cells by literal text or regex across one sheet or the whole workbook.                                                                                                                                          |
| `excel_add_sheet`            | Add a new empty sheet to the workbook session.                                                                                                                                                                       |
| `excel_delete_sheet`         | Delete a sheet from the workbook session. Cannot delete the only sheet.                                                                                                                                              |
| `excel_rename_sheet`         | Rename a sheet in the workbook session.                                                                                                                                                                              |
| `excel_copy_sheet`           | Duplicate a sheet within the same workbook while preserving state, printing, tables, drawings, rich cell semantics, and unsupported worksheet relationships with independently rebased target parts.                 |
| `excel_copy_sheet_to`        | Copy a sheet between loaded workbooks with collision-safe CodeName/table/drawing/passthrough-part rebasing.                                                                                                          |
| `excel_move_sheet`           | Move a sheet to a new position within the workbook.                                                                                                                                                                  |
| `excel_clone_rows`           | Deep-clone one or more rows and return them as a JSON array WITHOUT inserting.                                                                                                                                       |
| `excel_copy_row`             | Deep-copy a row, including rich-text runs and style/XF semantics, and insert it at a new position.                                                                                                                   |
| `excel_insert_rows`          | Insert rows at one or more positions in a single call.                                                                                                                                                               |
| `excel_insert_column`        | Insert a new empty column after the given column index.                                                                                                                                                              |
| `excel_edit_cells`           | Edit cell values across one or more rows — styles are preserved.                                                                                                                                                     |
| `excel_delete_rows`          | Delete one or more rows by index list or by a contiguous range.                                                                                                                                                      |
| `excel_clear_range`          | Clear values and/or styles from a rectangular cell range.                                                                                                                                                            |
| `excel_copy_column`          | Deep-copy a column and insert it after a given column index, preserving rich text and style/XF semantics.                                                                                                            |
| `excel_delete_column`        | Delete a column from all rows in a sheet.                                                                                                                                                                            |
| `excel_merge_cells`          | Merge a rectangular range of cells, or unmerge a merged region.                                                                                                                                                      |
| `excel_set_style`            | Set style properties on a single cell or a rectangular range.                                                                                                                                                        |
| `excel_set_font_color`       | Set font color on a single cell or range.                                                                                                                                                                            |
| `excel_set_strike`           | Enable or disable strikethrough on a single cell or range.                                                                                                                                                           |
| `excel_set_borders`          | Set or remove borders on a rectangular cell range.                                                                                                                                                                   |
| `excel_set_dimension`        | Set the height of a row or the width of a column.                                                                                                                                                                    |
| `excel_set_row_height`       | Set height for one or more rows in a single call.                                                                                                                                                                    |
| `excel_set_column_width`     | Set width for one or more columns in a single call.                                                                                                                                                                  |
| `excel_autofit_cols`         | Estimate and set column widths based on content length (heuristic approximation).                                                                                                                                    |
| `excel_freeze_panes`         | Freeze rows above `row` and/or columns to the left of `col`.                                                                                                                                                         |
| `excel_set_data_validation`  | Add a dropdown list validation to a cell range.                                                                                                                                                                      |
| `excel_find_rows`            | Find all rows where a column cell matches a value or regex pattern.                                                                                                                                                  |
| `excel_fill_column`          | Fill a column range with a constant value or an auto-incrementing sequence.                                                                                                                                          |
| `excel_fill_rows`            | Clone a template row N times and insert all copies in one call.                                                                                                                                                      |
| `excel_get_session_status`   | Return lightweight busy/operation and dirty-count status without reading or mutating workbook data.                                                                                                                  |

## Semantic and Expert Tools

The following public APIs cover lossless cell semantics, workbook metadata, structured objects, and expert OOXML package transactions. They mutate the in-memory session unless their description says otherwise.

| Tools | Description |
| --- | --- |
| `excel_create_workbook` | Create a new `.xlsx`/`.xltx` session or a macro/template-backed `.xlsm`/`.xltm` session without requiring a blank fixture. |
| `excel_get_rich_text`, `excel_edit_rich_text` | Read ordered rich-text runs and edit runs, Unicode substrings, lines, whitespace, and phonetic metadata without formatting the whole cell. |
| `excel_set_formula` | Set normal/shared/array/data-table formulas with explicit cached-value state and cache policy. |
| `excel_add_defined_name`, `excel_update_defined_name`, `excel_delete_defined_name` | Manage workbook- or worksheet-scoped defined names and their extended metadata. |
| `excel_set_auto_filter` | Patch or replace filter columns and sort state without dropping sibling criteria. |
| `excel_add_table`, `excel_update_table`, `excel_delete_table` | Manage table structure, columns, formulas, totals metadata, filters, and style flags. |
| `excel_set_hyperlink`, `excel_remove_hyperlink` | Manage external/internal hyperlinks while preserving display text and tooltip metadata on partial updates. |
| `excel_set_comment`, `excel_remove_comment`, `excel_set_ignored_errors` | Manage legacy comments and ignored-error rules; unrelated VML/comment shapes are preserved. |
| `excel_add_named_style`, `excel_update_named_style`, `excel_delete_named_style`, `excel_set_cell_style_semantics` | Manage named styles and expert cell XF flags, including dry-run inspection. |
| `excel_get_workbook_semantics`, `excel_set_calculation_properties`, `excel_set_workbook_properties`, `excel_set_document_properties`, `excel_set_workbook_protection` | Inspect or patch calculation, workbook, document-property, and workbook-protection semantics. |
| `excel_get_workbook_views`, `excel_set_workbook_views` | Inspect or patch the complete workbook-view list. |
| `excel_get_sheet_semantics`, `excel_set_sheet_state`, `excel_set_sheet_properties`, `excel_get_sheet_views`, `excel_set_sheet_views`, `excel_set_row_properties`, `excel_set_sheet_protection` | Inspect or patch worksheet state, properties, views, selections, row metadata, and legacy/modern protection hashes. |
| `excel_set_page_setup`, `excel_set_print_options`, `excel_set_header_footer`, `excel_set_page_breaks`, `excel_set_print_area`, `excel_set_print_titles` | Patch printing and page-layout semantics. |
| `excel_set_protected_ranges` | Add, update, or delete worksheet editable ranges without implicitly toggling sheet protection. |
| `excel_add_image`, `excel_add_chart`, `excel_add_shape` | Create DrawingML objects with anchors and package relationships; shapes accept the common rich-text run model. |
| `excel_list_package_parts`, `excel_read_package_part` | Inspect bounded package part content, hashes, content types, and relationships. |
| `excel_upsert_package_part`, `excel_delete_package_part`, `excel_set_package_relationships`, `excel_set_package_content_types`, `excel_apply_package_transaction` | Atomically create/update/delete OOXML parts and their graph metadata with traversal, XML, relationship-ID, and dangling-reference validation. |

## Heavy-Operation Cancellation

`excel_load`, `excel_verify_preservation`, `excel_save`, and `excel_save_as_copy` run their CPU- and I/O-heavy phases in private named subprocesses. The parent validates bounded request metadata, polls atomic status/result artifacts, verifies result size and SHA-256, terminates the complete process tree when required, and removes the worker workspace before returning. Worker stdin/stdout are disconnected from MCP stdio, and one heavy worker is allowed by default.

Cancellable load serializes the workbook only inside the child process and writes the full session model to a private atomic UTF-8 JSON artifact instead of a Pipe. Tagged values preserve `datetime`, `date`, `time`, and `bytes`; the parent validates size, SHA-256, UTF-8, JSON, and the workbook model before publishing a session. JSON decoding and session import run outside the MCP event loop. Timeout, cancellation, worker error, invalid data, or artifact-integrity failure creates no new session and leaves no worker workspace/result artifact.

Transactional save serializes the trusted in-memory session to a private `input.pkl` artifact with a configured size limit plus SHA-256 metadata. The worker validates that artifact, reconstructs into a unique staging workbook in the destination directory, and optionally verifies preservation while the destination remains unchanged. The parent creates the two-day backup only at the final commit boundary, confirms its SHA-256, and uses `os.replace()` for the same-filesystem commit. Failed verification, timeout, cancellation, worker failure, backup failure, or replacement failure leaves the destination and session baseline/dirty state unchanged.

A per-session busy registry blocks mutation, reload, close, replacement load, and package edits throughout save, verify, and commit. `excel_get_session_status` remains available as a lightweight read-only status call. Direct Python callers retain synchronous `excel_save` and `excel_save_as_copy` functions; MCP registration uses async wrappers under the same public tool names.

On Windows source runs, the worker bypasses the virtual-environment redirector but explicitly removes foreign global/user `site-packages` entries before importing the server. This keeps the worker on the same pinned dependencies as the parent (`mcp==1.28.1`, `openpyxl==3.1.5`). Frozen builds embed server, commit, and library metadata so preservation reports keep non-null runtime provenance outside a Git checkout.

The pinned MCP SDK responds to cancellation before a normal tool handler can finish cleanup. The Excel server therefore defers cancellation responses only for its named heavy tools until worker termination and workspace/staging cleanup complete; returned errors include `worker_stopped`, `workspace_removed`, `failure_paths_removed`, and `staging_removed`. Frozen stdio uses owned UTF-8 wrappers that detach from the underlying handles at shutdown, preventing the PyInstaller bootloader from flushing a handle already closed by the MCP transport.

`timeout_seconds` applies to slot waiting, worker startup, execution, artifact validation, and cleanup. Its accepted range is `0.1` through `3600` seconds. Operation-specific configuration uses `EXCEL_MCP_LOAD_TIMEOUT_SECONDS` for load, `EXCEL_MCP_VERIFY_TIMEOUT_SECONDS` for verify, and `EXCEL_MCP_SAVE_TIMEOUT_SECONDS` for save, then falls back to `EXCEL_MCP_HEAVY_TIMEOUT_SECONDS` and the built-in default `240`.

Additional startup configuration:

| Variable | Default | Accepted range | Purpose |
| --- | ---: | ---: | --- |
| `EXCEL_MCP_MAX_HEAVY_WORKERS` | `1` | `1`-`8` | Maximum concurrent heavy-operation workers; read when the server starts. |
| `EXCEL_MCP_MAX_RESULT_BYTES` | `67108864` | `1048576`-`2147483648` | Maximum accepted non-load worker result artifact size. |
| `EXCEL_MCP_MAX_LOAD_RESULT_BYTES` | `536870912` | `1048576`-`2147483648` | Maximum accepted load-session JSON artifact size. |
| `EXCEL_MCP_MAX_INPUT_BYTES` | `536870912` | `1048576`-`2147483648` | Maximum trusted parent-generated input artifact size. |

Timeout and cancellation failures use stable codes such as `HEAVY_OPERATION_TIMEOUT` and `HEAVY_OPERATION_CANCELLED`, and report cleanup completion only after the worker/process tree has stopped. Unexpected bootstrap, worker-exit, artifact-integrity, backup, replacement, and cleanup failures use separate structured codes.
## Notes

- `convert_to_markdown` is read-only, does not require `excel_load`, and is broader than the edit tools because it does not save back to the source file.
- Editing is session-based: load first, mutate the in-memory workbook, then save.
- Prefer `excel_save_as_copy` for complex workbooks until the output has been validated and opened successfully.
- Macro/template OOXML formats (`.xlsm`, `.xltm`, `.xltx`) are preserved best-effort during session/edit/save. Legacy/binary `.xls` and `.xlsb` are not handled by the edit engine; convert them to OOXML first for editing.
- Use `excel_validate_workbook` before and after risky workflows to confirm the package remains structurally valid.
- Every `excel_save` creates a managed pre-save backup retained for two days; set `DOCLOUPE_EXCEL_BACKUP_DIR` to override its temporary directory.
- Pass `report_format="json"` to `excel_save` or `excel_save_as_copy` for backup metadata, package-signature status, warnings, requested/actual semantic paths, and verifier instructions; add `verify_preservation=true` to run the semantic verifier inline. Signed packages report byte preservation only: intentional edits require re-signing, and changed/missing `_xmlsignatures/*` parts produce a critical warning; cryptographic validity is never claimed.
- Use `excel_verify_preservation` separately when you need full before/after detail for values, formulas, rich text, cell presence, style semantics, workbook/sheet settings, printing, filtering, protection, tables, properties, and advanced OOXML parts. Reports include file hashes, package validity/loadability, runtime/commit metadata, bounded previews, and all six classifications; pass `fixture_gap_paths` or `verifier_gap_paths` only for known evidence limitations. Omit `before_path` only when the latest backup is the intended reference.
- Use `excel_diff_package` for lower-level diagnostics of package parts that were added, removed, or semantically modified.
- Drawing support is intentionally conservative: tools can create images, charts, and simple shapes; inspect captured DrawingML objects; and update simple text/style properties. Complex layout/grouping plus VML/comment shapes, pivot tables, slicers, and external links remain preservation-only or require the expert package transaction API.
- Row/column insert, delete, and copy operations shift formulas, defined names, filters, tables, validations, conditional formatting, hyperlinks, sheet selections, page breaks, and DrawingML anchors.
- Row and column indices exposed by editing tools are 0-based unless a tool docstring states otherwise.
- `excel_capture` depends on LibreOffice being installed and available on the host.

