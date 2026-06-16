# excel-tools

Round-trip Excel workbook editor with one-shot Markdown export plus session-based reads, sheet operations, rows/columns, styles, merges, data validation, images, and save/reload support.

Use this server when an agent needs Excel content without dumping an entire workbook into context. The tools can inspect only the needed sheets, rows, columns, or cells, which helps reduce token usage and quota consumption.

## Status

Implemented and included in the completed DocForge MCP server set.

## Input

- `convert_to_markdown` is read-only and can convert Excel-family files readable by the converter into Markdown without creating a session.
- All session/edit/save tools are `.xlsx`-only; macro/binary formats are rejected there to avoid silent data loss.
- For non-Excel document types, use the matching DocForge MCP server: Markdown, PDF, DOCX, PPTX, CSV, HTML, text, JSON, or JSONL.

## Recommended Workflow

- Start with `excel_get_info` for sheet dimensions.
- Use `convert_to_markdown` when you only need a Markdown export and do not need a session.
- Prefer targeted reads such as `excel_get_rows`, `excel_get_cell`, and `excel_get_column` to avoid spending tokens on irrelevant workbook content.
- Call `excel_load` and keep the returned `session_key`.
- Inspect with `excel_to_markdown`, `excel_get_rows`, or cell/column tools.
- Mutate the session, then call `excel_save` to write an `.xlsx`.

## Tools

| Tool | Description |
| --- | --- |
| `convert_to_markdown` | Convert an Excel-family file to Markdown in one call without creating a session. |
| `excel_get_info` | Return summary info about an Excel file: sheet names, row and column counts. |
| `excel_load` | Load an Excel file into the server session cache and return a session_key. |
| `excel_save` | Reconstruct an Excel file from session data and write it to disk. |
| `excel_reload` | Reload session data from disk, discarding any unsaved in-memory changes. |
| `excel_close` | Remove a session from the server cache to free memory. |
| `excel_to_markdown` | Export session data as Markdown tables annotated with 0-based row/column indices. |
| `excel_capture` | Render a sheet as a PNG image using LibreOffice. |
| `excel_extract_images` | Extract all embedded images from a sheet and save them to a directory. |
| `excel_get_rows` | Get a range of rows from session data as JSON. |
| `excel_get_cell` | Get full metadata of a single cell. |
| `excel_get_column` | Get all cells in a column as JSON. |
| `excel_add_sheet` | Add a new empty sheet to the workbook session. |
| `excel_delete_sheet` | Delete a sheet from the workbook session. Cannot delete the only sheet. |
| `excel_rename_sheet` | Rename a sheet in the workbook session. |
| `excel_copy_sheet` | Duplicate a sheet within the same workbook session. |
| `excel_copy_sheet_to` | Copy a sheet from one loaded workbook session into another. |
| `excel_move_sheet` | Move a sheet to a new position within the workbook. |
| `excel_clone_rows` | Deep-clone one or more rows and return them as a JSON array WITHOUT inserting. |
| `excel_copy_row` | Clone a row and insert the copy immediately at a new position (one step). |
| `excel_insert_rows` | Insert rows at one or more positions in a single call. |
| `excel_insert_column` | Insert a new empty column after the given column index. |
| `excel_edit_cells` | Edit cell values across one or more rows — styles are preserved. |
| `excel_delete_rows` | Delete one or more rows by index list or by a contiguous range. |
| `excel_clear_range` | Clear values and/or styles from a rectangular cell range. |
| `excel_copy_column` | Copy a column and insert it after a given column index, preserving all styles. |
| `excel_delete_column` | Delete a column from all rows in a sheet. |
| `excel_merge_cells` | Merge a rectangular range of cells, or unmerge a merged region. |
| `excel_set_style` | Set style properties on a single cell or a rectangular range. |
| `excel_set_borders` | Set or remove borders on a rectangular cell range. |
| `excel_set_dimension` | Set the height of a row or the width of a column. |
| `excel_set_row_height` | Set height for one or more rows in a single call. |
| `excel_set_column_width` | Set width for one or more columns in a single call. |
| `excel_autofit_cols` | Estimate and set column widths based on content length (heuristic approximation). |
| `excel_freeze_panes` | Freeze rows above `row` and/or columns to the left of `col`. |
| `excel_set_data_validation` | Add a dropdown list validation to a cell range. |
| `excel_find_rows` | Find all rows where a column cell matches a value or regex pattern. |
| `excel_fill_column` | Fill a column range with a constant value or an auto-incrementing sequence. |
| `excel_fill_rows` | Clone a template row N times and insert all copies in one call. |


## Notes

- `convert_to_markdown` is read-only, does not require `excel_load`, and is broader than the edit tools because it does not save back to the source file.
- Editing is session-based: load first, mutate the in-memory workbook, then save.
- Unsupported macro/binary formats (`.xlsm`, `.xltm`, `.xlsb`, `.xls`) are rejected by session/edit/save tools to avoid silent data loss.
- Row and column indices exposed by editing tools are 0-based unless a tool docstring states otherwise.
- `excel_capture` depends on LibreOffice being installed and available on the host.
