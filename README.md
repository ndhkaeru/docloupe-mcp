# DocForge MCP

> **Let coding agents read, convert, and edit common document formats through focused local MCP servers.**

`docforge-mcp` is a local-first [Model Context Protocol](https://modelcontextprotocol.io) toolkit for document workflows. It ships separate `stdio` MCP servers for Markdown, Excel, PDF, DOCX, PPTX, CSV, HTML, plain text, and JSON/JSONL so agents can use the right tool without loading whole files or ad-hoc conversion code into context.

Instead of pasting entire documents into chat, agents can call targeted tools that return only the needed section, table, cell range, or converted Markdown. This usually means smaller context payloads, less token usage, and lower quota consumption while keeping edits more precise.

Everything runs on your machine. There is no hosted service, no remote indexing, and no telemetry.

---

## Highlights

- **Markdown-native editing**: outline headings, read/replace/move sections, edit tables, manage diagrams/code blocks, update TOCs, inspect links/images, and work with frontmatter.
- **Excel round-trip editing**: load `.xlsx` files into sessions, inspect rows/cells/styles, edit sheets/rows/columns, validate saves before replace, and preserve advanced workbook parts best-effort.
- **Broad file support**: use dedicated tools for Markdown, Excel, PDF, DOCX, PPTX, CSV, HTML, plain text, JSON, and JSONL workflows.
- **Token/quota efficient**: query just the relevant outline, section, table, cell, or converted text instead of sending full files to the model.
- **Document conversion**: convert PDF, DOCX, PPTX, CSV, HTML, JSON/JSONL, text, and Excel files into Markdown/plain text that agents can reason over.
- **Small server boundaries**: enable only the MCP servers you need instead of one large all-purpose binary.
- **Local-only by default**: files remain on the machine running the MCP server.
- **Release binaries**: native artifacts are published for `windows-x64`, `linux-x64`, `macos-x64`, and `macos-arm64`.

---

## Quick Start

### 1. Get the binaries

Download the latest binaries for your platform from the [Releases page](https://github.com/ndhkaeru/docforge-mcp/releases). Supported platform keys are `windows-x64`, `linux-x64`, `macos-x64`, and `macos-arm64`.

Each server is published as a separate executable:

| Binary | Use it for |
| --- | --- |
| `docforge-mcp-md-tools-<platform>[.exe]` | Structure-aware Markdown reads and edits. |
| `docforge-mcp-excel-tools-<platform>[.exe]` | Excel-family Markdown export plus `.xlsx` workbook inspection and editing. |
| `docforge-mcp-pdf-tools-<platform>[.exe]` | PDF-to-Markdown text extraction. |
| `docforge-mcp-docx-tools-<platform>[.exe]` | DOCX-to-Markdown conversion. |
| `docforge-mcp-pptx-tools-<platform>[.exe]` | PPTX-to-Markdown conversion. |
| `docforge-mcp-csv-tools-<platform>[.exe]` | CSV-to-Markdown tables. |
| `docforge-mcp-html-tools-<platform>[.exe]` | HTML-to-Markdown conversion. |
| `docforge-mcp-text-tools-<platform>[.exe]` | Plain text and Markdown passthrough reads. |
| `docforge-mcp-json-tools-<platform>[.exe]` | JSON/JSONL passthrough reads. |

A `*-sha256sums.txt` file is attached to each release for integrity checks.

### 2. Add a server to your MCP client

#### Option A: npm/npx

Use the npm launcher and pass the server name as the first argument:

```json
{
  "mcpServers": {
    "docforge-excel": {
      "command": "npx",
      "args": ["-y", "@ndhkaeru/docforge-mcp@latest", "excel"]
    },
    "docforge-md": {
      "command": "npx",
      "args": ["-y", "@ndhkaeru/docforge-mcp@latest", "md"]
    }
  }
}
```

Supported npm server arguments: `excel`, `md`, `pdf`, `docx`, `pptx`, `csv`, `html`, `text`, and `json`.

#### Option B: downloaded native binaries

Register the desired binary as a `stdio` server. Replace the path with your downloaded platform binary; Windows artifacts end in `.exe`, Linux/macOS artifacts do not.

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
| Risk losing workbook styles or corrupting saves while editing | Excel tools validate `.xlsx` output before replace and preserve styles, merges, dimensions, validations, comments, drawings, charts, and images where supported |
| Ask an agent to parse DOCX/PDF/PPTX directly | Conversion servers emit Markdown/plain text first |

The result: smaller tool responses, lower token/quota usage, safer edits, and document workflows that are easier to audit.

---

## Servers and Tools

### `md-tools`

Markdown structure editor for `.md` and `.markdown` files.

| Group | Tools |
| --- | --- |
| Outline/search/read | `markdown_outline`, `md_get_document_map`, `read_markdown_section`, `md_read_range`, `md_read_near`, `md_search`, `md_stats` |
| Section edits | `replace_markdown_section`, `md_insert_section`, `md_delete_section`, `md_append_to_section`, `md_patch_lines`, `md_move_section`, `md_set_heading_level`, `md_normalize_headings`, `md_rename_heading`, `md_replace_text` |
| Tables | `md_list_tables`, `md_read_table`, `md_format_table`, `md_edit_table`, `md_insert_table`, `md_table_export` |
| Diagrams/code | `md_list_diagrams`, `md_read_diagram`, `md_insert_diagram`, `md_replace_diagram`, `md_validate_diagram`, `md_render_diagram`, `md_extract_code_blocks`, `md_replace_code_block`, `md_insert_code_block`, `md_tangle` |
| Links/TOC/metadata | `md_list_links`, `md_rewrite_links`, `md_validate_links`, `md_check_internal_links`, `md_list_images`, `md_get_anchor`, `md_update_toc`, `md_frontmatter` |
| File operations | `md_split`, `md_merge`, `md_to_html` |

### `excel-tools`

One-shot Markdown export for Excel-family files plus session-based `.xlsx` workbook editing.

| Group | Tools |
| --- | --- |
| One-shot conversion | `convert_to_markdown` for read-only Markdown export without `excel_load`, `excel_get_workbook_summary`, `excel_get_sheet_preview` |
| Session lifecycle | `excel_get_info`, `excel_load`, `excel_save`, `excel_save_as_copy`, `excel_validate_workbook`, `excel_diff_package`, `excel_reload`, `excel_close` |
| Reading/export | `excel_to_markdown`, `excel_to_markdown_range`, `excel_read_range`, `excel_find_cells`, `excel_list_tables`, `excel_list_defined_names`, `excel_get_rows`, `excel_get_cell`, `excel_get_column`, `excel_capture`, `excel_extract_images`, `excel_get_shapes` |
| Sheets | `excel_add_sheet`, `excel_delete_sheet`, `excel_rename_sheet`, `excel_copy_sheet`, `excel_copy_sheet_to`, `excel_move_sheet` |
| Rows/cells | `excel_clone_rows`, `excel_copy_row`, `excel_insert_rows`, `excel_edit_cells`, `excel_delete_rows`, `excel_clear_range`, `excel_find_rows`, `excel_fill_rows` |
| Columns | `excel_insert_column`, `excel_copy_column`, `excel_delete_column`, `excel_fill_column` |
| Layout/style | `excel_merge_cells`, `excel_set_style`, `excel_set_font_color`, `excel_set_strike`, `excel_set_borders`, `excel_set_dimension`, `excel_set_row_height`, `excel_set_column_width`, `excel_autofit_cols`, `excel_freeze_panes`, `excel_set_data_validation`, `excel_update_shape_text`, `excel_set_shape_style` |

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
command = "/opt/docforge/docforge-mcp-md-tools-linux-x64"
args = []
enabled = true

[mcp_servers.excel-tools]
command = "/opt/docforge/docforge-mcp-excel-tools-linux-x64"
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

```bash
python -m venv .venv-build
.venv-build/bin/python -m pip install -U pip
.venv-build/bin/python -m pip install pytest pyinstaller "mcp[cli]" openpyxl pillow pdfminer.six pdfplumber mammoth python-pptx markdownify beautifulsoup4 lxml defusedxml charset-normalizer
```

```powershell
python -m venv .venv-build
.\.venv-build\Scripts\python.exe -m pip install -U pip
.\.venv-build\Scripts\python.exe -m pip install pytest pyinstaller "mcp[cli]" openpyxl pillow pdfminer.six pdfplumber mammoth python-pptx markdownify beautifulsoup4 lxml defusedxml charset-normalizer
```

### Run tests

```bash
.venv-build/bin/python -m pytest -q
```

```powershell
.\.venv-build\Scripts\python.exe -m pytest -q
```

### Build locally

On Windows, use the root helper:

```powershell
.uild.ps1              # all servers
.uild.ps1 -Only excel  # one server smoke build
```

On Linux/macOS, use the same PyInstaller shape as the release workflow:

```bash
.venv-build/bin/python -m PyInstaller --clean --onefile --name text-tools --paths servers/text --specpath build --workpath build --distpath dist servers/text/main.py
```

Build output goes to `dist/*-tools[.exe]` and is ignored by git.

---

## Shipping

Releases are distributed as GitHub Release artifacts. A semver tag such as `v0.1.1` runs `.github/workflows/release.yml`, builds every native server executable for Windows, Linux, and macOS, and uploads checksums.

Release checklist:

1. Update `server.json` version fields when cutting a versioned release.
2. Run the setup commands above and the platform venv Python with `-m pytest -q`.
3. Build at least the affected server locally, for example `./build.ps1 -Only excel` on Windows or the PyInstaller command above on Linux/macOS.
4. Push `main` and confirm the `Push` workflow succeeds.
5. Push a semver tag such as `v0.1.1` and confirm the `Release` workflow uploads every executable.
6. Smoke test downloaded artifacts on real files when round-trip behavior changed.

### MCP Registry Metadata

`server.json` records the repository identity and release version in the MCP Registry metadata shape. Keep its top-level `version` and package entry version in sync with GitHub Release tags.

### Tool Description Style

Tool descriptions are written for agent routing: they state when to use the tool, which arguments scope output, which operations are best-effort, and which save/validation steps protect user files.

---

## Notes and Limits

- Excel `convert_to_markdown` is read-only and can handle Excel-family files readable by the converter. Excel session/edit/save tools support OOXML workbooks (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) and reject unsafe macro/non-macro extension mismatches.
- Advanced Excel objects such as DrawingML, macros, external links, pivot tables, slicers, and VML are preserved best-effort; validation warns when advanced parts become orphaned.
- Use limit parameters such as `sheet_name`, `range_ref`, `max_rows`, `max_cols`, `preview`, `max_chars`, and `include_body=false` when available to keep tool responses small.
- `excel_capture` requires LibreOffice on the machine running the MCP server.
- Markdown diagram validation/rendering depends on optional external CLIs such as Mermaid CLI.
- PDF conversion is best-effort text extraction, not pixel-perfect layout reconstruction.
- Release binaries are native executables.

---

## License

Licensed under the [Apache License 2.0](LICENSE).

