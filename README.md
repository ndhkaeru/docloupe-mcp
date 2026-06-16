# DocForge MCP

![Python](https://img.shields.io/badge/Python-3.12-blue?logo=python)
![MCP](https://img.shields.io/badge/MCP-stdio-blue)
![Platform](https://img.shields.io/badge/Windows-x64-lightgrey?logo=windows)
![Release](https://img.shields.io/github/v/release/ndhkaeru/docforge-mcp?include_prereleases)
![License](https://img.shields.io/badge/license-Apache--2.0-green)

> **Let coding agents read, convert, and edit common document formats through focused local MCP servers.**

`docforge-mcp` is a local-first [Model Context Protocol](https://modelcontextprotocol.io) toolkit for document workflows. It ships separate `stdio` MCP servers for Markdown, Excel, PDF, DOCX, PPTX, CSV, HTML, plain text, and JSON/JSONL so agents can use the right tool without loading whole files or ad-hoc conversion code into context.

Instead of pasting entire documents into chat, agents can call targeted tools that return only the needed section, table, cell range, or converted Markdown. This usually means smaller context payloads, less token usage, and lower quota consumption while keeping edits more precise.

Everything runs on your machine. There is no hosted service, no remote indexing, and no telemetry.

---

## Highlights

- **Markdown-native editing**: outline headings, read/replace/move sections, edit tables, manage diagrams/code blocks, update TOCs, inspect links/images, and work with frontmatter.
- **Excel round-trip editing**: load `.xlsx` files into sessions, inspect rows/cells/styles, edit sheets/rows/columns, preserve workbook structure, and save back to `.xlsx`.
- **Broad file support**: use dedicated tools for Markdown, Excel, PDF, DOCX, PPTX, CSV, HTML, plain text, JSON, and JSONL workflows.
- **Token/quota efficient**: query just the relevant outline, section, table, cell, or converted text instead of sending full files to the model.
- **Document conversion**: convert PDF, DOCX, PPTX, CSV, HTML, JSON/JSONL, text, and Excel files into Markdown/plain text that agents can reason over.
- **Small server boundaries**: enable only the MCP servers you need instead of one large all-purpose binary.
- **Local-only by default**: files remain on the machine running the MCP server.
- **Release binaries**: Windows x64 `.exe` artifacts are published for each server.

---

## Quick Start

### 1. Get the binaries

Download the latest Windows x64 binaries from the [Releases page](https://github.com/ndhkaeru/docforge-mcp/releases).

Each server is published as a separate executable:

| Binary | Use it for |
| --- | --- |
| `docforge-mcp-md-tools-windows-x64.exe` | Structure-aware Markdown reads and edits. |
| `docforge-mcp-excel-tools-windows-x64.exe` | Excel-family Markdown export plus `.xlsx` workbook inspection and editing. |
| `docforge-mcp-pdf-tools-windows-x64.exe` | PDF-to-Markdown text extraction. |
| `docforge-mcp-docx-tools-windows-x64.exe` | DOCX-to-Markdown conversion. |
| `docforge-mcp-pptx-tools-windows-x64.exe` | PPTX-to-Markdown conversion. |
| `docforge-mcp-csv-tools-windows-x64.exe` | CSV-to-Markdown tables. |
| `docforge-mcp-html-tools-windows-x64.exe` | HTML-to-Markdown conversion. |
| `docforge-mcp-text-tools-windows-x64.exe` | Plain text and Markdown passthrough reads. |
| `docforge-mcp-json-tools-windows-x64.exe` | JSON/JSONL passthrough reads. |

A `*-sha256sums.txt` file is attached to each release for integrity checks.

### 2. Add a server to your MCP client

Register the desired binary as a `stdio` server. Replace the path with your downloaded `.exe` path.

**OpenAI Codex CLI**

```powershell
codex mcp add docforge-md -- C:\path\to\docforge-mcp-md-tools-windows-x64.exe
codex mcp add docforge-excel -- C:\path\to\docforge-mcp-excel-tools-windows-x64.exe
```

**Claude Code**

```powershell
claude mcp add docforge-md -- C:\path\to\docforge-mcp-md-tools-windows-x64.exe
```

**Gemini CLI**

```powershell
gemini mcp add docforge-md -- C:\path\to\docforge-mcp-md-tools-windows-x64.exe
```

> Prefer editing config files directly, or using Cursor, Windsurf, Cline, Roo, or VS Code?
> See [Client Configuration](#client-configuration) below.

---

## Why DocForge MCP?

Agents waste context and become brittle when they rely on raw file dumps or one-off conversion scripts. DocForge MCP gives them bounded, format-aware tools:

| Without it | With `docforge-mcp` |
| --- | --- |
| Read a whole Markdown file to find one section | `markdown_outline` + `read_markdown_section` target exactly what is needed |
| Rewrite a Markdown table by hand | `md_read_table`, `md_format_table`, and `md_edit_table` operate on table structure |
| Dump an entire workbook into context | `excel_get_info`, `excel_get_rows`, and `excel_get_cell` inspect only the relevant sheet/range |
| Risk losing workbook styles while editing | Excel tools preserve styles, merges, dimensions, validations, comments, and images where supported |
| Ask an agent to parse DOCX/PDF/PPTX directly | Conversion servers emit Markdown/plain text first |

The result: smaller tool responses, lower token/quota usage, safer edits, and document workflows that are easier to audit.

---

## Servers and Tools

### `md-tools`

Markdown structure editor for `.md` and `.markdown` files.

| Group | Tools |
| --- | --- |
| Outline/search/read | `markdown_outline`, `read_markdown_section`, `md_search`, `md_stats` |
| Section edits | `replace_markdown_section`, `md_insert_section`, `md_delete_section`, `md_append_to_section`, `md_move_section`, `md_set_heading_level`, `md_rename_heading`, `md_replace_text` |
| Tables | `md_list_tables`, `md_read_table`, `md_format_table`, `md_edit_table`, `md_insert_table`, `md_table_export` |
| Diagrams/code | `md_list_diagrams`, `md_read_diagram`, `md_insert_diagram`, `md_replace_diagram`, `md_validate_diagram`, `md_render_diagram`, `md_extract_code_blocks`, `md_replace_code_block`, `md_insert_code_block`, `md_tangle` |
| Links/TOC/metadata | `md_list_links`, `md_rewrite_links`, `md_validate_links`, `md_list_images`, `md_get_anchor`, `md_update_toc`, `md_frontmatter` |
| File operations | `md_split`, `md_merge`, `md_to_html` |

### `excel-tools`

One-shot Markdown export for Excel-family files plus session-based `.xlsx` workbook editing.

| Group | Tools |
| --- | --- |
| One-shot conversion | `convert_to_markdown` for read-only Markdown export without `excel_load` |
| Session lifecycle | `excel_get_info`, `excel_load`, `excel_save`, `excel_reload`, `excel_close` |
| Reading/export | `excel_to_markdown`, `excel_get_rows`, `excel_get_cell`, `excel_get_column`, `excel_capture`, `excel_extract_images` |
| Sheets | `excel_add_sheet`, `excel_delete_sheet`, `excel_rename_sheet`, `excel_copy_sheet`, `excel_copy_sheet_to`, `excel_move_sheet` |
| Rows/cells | `excel_clone_rows`, `excel_copy_row`, `excel_insert_rows`, `excel_edit_cells`, `excel_delete_rows`, `excel_clear_range`, `excel_find_rows`, `excel_fill_rows` |
| Columns | `excel_insert_column`, `excel_copy_column`, `excel_delete_column`, `excel_fill_column` |
| Layout/style | `excel_merge_cells`, `excel_set_style`, `excel_set_borders`, `excel_set_dimension`, `excel_set_row_height`, `excel_set_column_width`, `excel_autofit_cols`, `excel_freeze_panes`, `excel_set_data_validation` |

### Conversion servers

Each conversion server exposes one focused `convert_to_markdown` tool.

| Server | Input | Output |
| --- | --- | --- |
| `pdf-tools` | `.pdf` | Best-effort Markdown text extracted from PDF pages. |
| `docx-tools` | `.docx` | Markdown text from Word document content. |
| `pptx-tools` | `.pptx` | Markdown text from slide content in presentation order. |
| `csv-tools` | `.csv` | Markdown pipe table. |
| `html-tools` | `.html`, `.htm` | Markdown text converted from HTML. |
| `text-tools` | `.txt`, `.md` | Raw text content. |
| `json-tools` | `.json`, `.jsonl` | Raw JSON/JSONL text content. |

---

## Example Tool Calls

Read a Markdown outline:

```json
{
  "name": "markdown_outline",
  "arguments": {
    "path": "C:\\docs\\guide.md"
  }
}
```

Append to a Markdown section:

```json
{
  "name": "md_append_to_section",
  "arguments": {
    "path": "C:\\docs\\guide.md",
    "heading_path": ["Guide", "Install"],
    "content": "New troubleshooting note.",
    "position": "end"
  }
}
```

Load and edit an Excel workbook:

```json
{
  "name": "excel_load",
  "arguments": {
    "uri": "C:\\workbooks\\scores.xlsx"
  }
}
```

```json
{
  "name": "excel_edit_cells",
  "arguments": {
    "session_key": "C:\\workbooks\\scores.xlsx",
    "sheet_name": "Scores",
    "edits": [
      { "row_index": 2, "edits": { "1": 25 } }
    ]
  }
}
```

Convert a PDF to Markdown:

```json
{
  "name": "convert_to_markdown",
  "arguments": {
    "file_path": "C:\\docs\\paper.pdf"
  }
}
```

---

## Client Configuration

Use one MCP entry per server binary you want enabled.

<details>
<summary><strong>OpenAI Codex CLI config</strong></summary>

```toml
[mcp_servers.md-tools]
command = "C:/tools/docforge-mcp-md-tools-windows-x64.exe"
args = []
enabled = true

[mcp_servers.excel-tools]
command = "C:/tools/docforge-mcp-excel-tools-windows-x64.exe"
args = []
enabled = true
```
</details>

<details>
<summary><strong>Claude Desktop / compatible JSON config</strong></summary>

```json
{
  "mcpServers": {
    "docforge-md": {
      "command": "C:\\tools\\docforge-mcp-md-tools-windows-x64.exe",
      "args": []
    },
    "docforge-excel": {
      "command": "C:\\tools\\docforge-mcp-excel-tools-windows-x64.exe",
      "args": []
    }
  }
}
```
</details>

---

## Development

### Set up a local environment

```powershell
python -m venv .venv-build
.\.venv-build\Scripts\python.exe -m pip install -U pip
.\.venv-build\Scripts\python.exe -m pip install pytest pyinstaller "mcp[cli]" openpyxl pillow pdfminer.six pdfplumber mammoth python-pptx markdownify beautifulsoup4 lxml defusedxml charset-normalizer
```

### Run tests

```powershell
.\.venv-build\Scripts\python.exe -m pytest -q
```

### Build locally

Prepare the complete source tree, then use the PyInstaller commands mirrored in `.github/workflows/release.yml` or your local build helper.

Build output goes to `dist\*-tools.exe` and is ignored by git.

---

## Release Process

The release workflow runs when a tag matching `v*.*.*` is pushed.

```powershell
git tag -a v1.0.0 -m "Release v1.0.0"
git push origin v1.0.0
```

The workflow verifies tests, builds Windows x64 executables, generates SHA256 checksums, and publishes a GitHub Release.

---

## Notes and Limits

- Excel `convert_to_markdown` is read-only and can handle Excel-family files readable by the converter; Excel session/edit/save tools are `.xlsx`-only to avoid silent data loss.
- `excel_capture` requires LibreOffice on the machine running the MCP server.
- Markdown diagram validation/rendering depends on optional external CLIs such as Mermaid CLI.
- PDF conversion is best-effort text extraction, not pixel-perfect layout reconstruction.
- Release binaries are Windows x64 executables.

---

## Repository Hygiene

Commit source, tests, docs, and workflows. Do not commit local environments, build outputs, release artifacts, logs, or local-only build helpers.

---

## License

Licensed under the [Apache License 2.0](LICENSE).
