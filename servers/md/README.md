# md-tools

Markdown structure editor for headings, sections, tables, diagrams, code blocks, links, TOC, frontmatter, split/merge, and lightweight HTML export.

## Status

Implemented and included in the completed DocLoupe MCP server set.

## Input

.md, .markdown

## Recommended Workflow

- Run `markdown_outline` or `md_get_document_map` before structural edits.
- Use `md_read_range` or `md_read_near` for targeted context instead of reading whole files.
- Use `preview`, `max_chars`, `max_lines`, `max_bytes`, or `include_body=false` on large reads to reduce token/quota usage.
- Use `heading_path` for duplicate headings.
- Re-read affected sections after edits.
- Use `text-tools` for arbitrary whole-file reads.

## Tools

| Tool | Description |
| --- | --- |
| `markdown_outline` | List Markdown headings with line numbers, hierarchy, and section spans. |
| `read_markdown_section` | Return a Markdown section as raw Markdown text by heading name or heading path; supports preview and size limits. |
| `md_read_range` | Read a precise 1-based inclusive Markdown line range. |
| `md_read_near` | Read Markdown context around a matching text/regex or around a line number. |
| `md_get_document_map` | Return a compact map of headings, tables, diagrams, code blocks, links, and images. |
| `replace_markdown_section` | Replace a Markdown section by heading name or heading path without exact text matching. |
| `md_patch_lines` | Replace a precise 1-based inclusive line range with Markdown text. |
| `md_search` | Search Markdown text and return line matches with heading context. |
| `md_insert_section` | Insert a new Markdown section at the document end or under a parent heading. |
| `md_delete_section` | Delete a Markdown section by heading name or heading path. |
| `md_append_to_section` | Append or prepend Markdown content inside a section body without overwriting it. |
| `md_rename_heading` | Rename a Markdown heading while preserving its level and section body. |
| `md_replace_text` | Find and replace literal or regex text in the whole document or one section. |
| `md_list_tables` | List Markdown pipe tables with line ranges, heading context, shape, and headers. |
| `md_read_table` | Read a Markdown pipe table as structured headers and rows; supports `include_body=false` and `max_rows`. |
| `md_format_table` | Pretty-print one Markdown pipe table with normalized spacing and alignment markers. |
| `md_edit_table` | Edit a Markdown table using op: set_cell, add_row, add_col, del_row, del_col, sort, set_align. |
| `md_insert_table` | Insert a Markdown table from structured headers and rows. |
| `md_table_export` | Export one Markdown table to CSV or JSON text. |
| `md_list_diagrams` | List fenced diagram blocks such as mermaid, plantuml, and dot. |
| `md_read_diagram` | Read the source of a fenced diagram block; supports metadata-only and preview reads. |
| `md_insert_diagram` | Insert a fenced diagram block. |
| `md_replace_diagram` | Replace the source inside one fenced diagram block. |
| `md_extract_code_blocks` | List fenced code blocks and optionally return source with preview/size limits. |
| `md_replace_code_block` | Replace the source inside a fenced code block while keeping its fence and info string. |
| `md_insert_code_block` | Insert a fenced code block with an optional info string. |
| `md_list_links` | List inline and reference Markdown links with line numbers. |
| `md_list_images` | List Markdown images with alt text, source, and line numbers. |
| `md_get_anchor` | Return the GitHub-compatible anchor slug for a heading title. |
| `md_update_toc` | Generate or refresh a Markdown table of contents between marker comments. |
| `md_stats` | Return simple Markdown document statistics. |
| `md_set_heading_level` | Set a heading level and shift every heading in its subtree by the same delta. |
| `md_normalize_headings` | Normalize heading hierarchy so levels do not jump unexpectedly. |
| `md_move_section` | Move a section before/after another heading, or under another heading as a child. |
| `md_check_internal_links` | Validate local file links and same-file heading anchors without remote checks. |
| `md_rewrite_links` | Bulk rewrite Markdown link and image targets. |
| `md_validate_links` | Validate local file links and same-file heading anchors; remote links are counted as skipped. |
| `md_frontmatter` | Read or write simple YAML frontmatter using op: read, replace, set, delete. |
| `md_tangle` | Extract fenced code blocks to files for lightweight literate programming. |
| `md_split` | Split a Markdown document into files at headings of a chosen level. |
| `md_merge` | Merge multiple Markdown files, optionally offsetting heading levels. |
| `md_to_html` | Render a small Markdown subset to HTML without external dependencies. |
| `md_runtime_capabilities` | Report active diagram backends and preservation capabilities. |
| `md_validate_diagram` | Validate supported Mermaid and DOT syntax with the bundled renderer, or an optional host CLI. |
| `md_render_diagram` | Render supported Mermaid or DOT diagrams to SVG; Mermaid PNG/PDF needs host `mmdc`. |


## Notes

- ATX headings (`#` through `######`) are parsed; fenced code blocks are ignored when detecting headings.
- GitHub-compatible heading slugs keep Unicode letters, including Vietnamese text.
- New release builds bundle the Mermaid/DOT SVG renderer; source checkouts require building it or installing a compatible CLI. PlantUML validation still requires a host CLI. Use `md_runtime_capabilities` to inspect the active backend.
- To enable it from source, run `cargo build --locked --release --manifest-path servers/md/diagram_renderer/Cargo.toml --target-dir build/diagram-target` at the repository root. Set `DOCLOUPE_DIAGRAM_RENDERER` to an alternate renderer binary path if needed.
- For raw arbitrary file reads, prefer `text-tools`.
