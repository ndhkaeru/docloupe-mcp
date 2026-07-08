## DocForge MCP

A Model Context Protocol (MCP) package that provides local document workflow tools for Markdown, Excel, PDF, DOCX, PPTX, CSV, HTML, text, and JSON files.

DocForge MCP lets LLMs read, convert, inspect, and edit documents through focused stdio MCP servers instead of dumping entire files into model context.

### DocForge MCP vs one-off conversion scripts

This package provides MCP interfaces into DocForge's document tools.

- **One-off scripts** are useful when a coding agent only needs a single deterministic conversion command. They are lightweight, explicit, and easy to run in a shell.
- **MCP** is useful when an agent benefits from persistent sessions, structured reads, bounded edits, and iterative inspection of document state, such as editing Excel workbooks, manipulating Markdown sections/tables, or converting several file types through consistent tools.

### Key Features

- **Focused servers**. Enable only the document families you need: `excel`, `md`, `pdf`, `docx`, `pptx`, `csv`, `html`, `text`, or `json`.
- **LLM-friendly output**. Tools expose previews, ranges, summaries, and Markdown conversions instead of huge raw dumps.
- **Safer document edits**. Excel and Markdown tools edit structured content and validate save flows where supported.
- **Native binaries via npm**. The npm launcher downloads and runs platform-specific binaries from GitHub Releases.

### Requirements

- Node.js 18 or newer
- An MCP client such as VS Code, Cursor, Windsurf, Claude Desktop, Claude Code, Codex, Cline, Goose, Gemini CLI, Junie, Kiro, LM Studio, or another stdio MCP client
- One supported platform:
  - `win32-x64`
  - `linux-x64`
  - `darwin-x64`
  - `darwin-arm64`

### Getting started

First, install the DocForge MCP server with your client.

**Standard config** works in most tools:

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

Run one server from a terminal for a smoke test:

```bash
npx -y @ndhkaeru/docforge-mcp@latest excel
```

The command starts an MCP stdio server and waits for MCP client input.

<details>
<summary>Amp</summary>

Add via the Amp VS Code extension settings screen or by updating your settings JSON:

```json
"amp.mcpServers": {
  "docforge-excel": {
    "command": "npx",
    "args": ["-y", "@ndhkaeru/docforge-mcp@latest", "excel"]
  }
}
```

**Amp CLI setup:**

```bash
amp mcp add docforge-excel -- npx -y @ndhkaeru/docforge-mcp@latest excel
```

</details>

<details>
<summary>Antigravity</summary>

Add via Antigravity settings or by updating your MCP configuration file:

```json
{
  "mcpServers": {
    "docforge-excel": {
      "command": "npx",
      "args": ["-y", "@ndhkaeru/docforge-mcp@latest", "excel"]
    }
  }
}
```

</details>

<details>
<summary>Claude Code</summary>

Use the Claude Code CLI to add a DocForge MCP server:

```bash
claude mcp add docforge-excel -- npx -y @ndhkaeru/docforge-mcp@latest excel
claude mcp add docforge-md -- npx -y @ndhkaeru/docforge-mcp@latest md
```

</details>

<details>
<summary>Claude Desktop</summary>

Follow the MCP install guide and use the standard config above. Example:

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

</details>

<details>
<summary>Cline</summary>

Add the following to `cline_mcp_settings.json`:

```json
{
  "mcpServers": {
    "docforge-excel": {
      "type": "stdio",
      "command": "npx",
      "timeout": 60,
      "args": ["-y", "@ndhkaeru/docforge-mcp@latest", "excel"],
      "disabled": false
    }
  }
}
```

</details>

<details>
<summary>Codex</summary>

Use the Codex CLI:

```bash
codex mcp add docforge-excel -- npx -y @ndhkaeru/docforge-mcp@latest excel
codex mcp add docforge-md -- npx -y @ndhkaeru/docforge-mcp@latest md
```

Alternatively, create or edit `~/.codex/config.toml`:

```toml
[mcp_servers.docforge-excel]
command = "npx"
args = ["-y", "@ndhkaeru/docforge-mcp@latest", "excel"]

[mcp_servers.docforge-md]
command = "npx"
args = ["-y", "@ndhkaeru/docforge-mcp@latest", "md"]
```

</details>

<details>
<summary>Copilot</summary>

Use Copilot's MCP add flow or create/edit `~/.copilot/mcp-config.json`:

```json
{
  "mcpServers": {
    "docforge-excel": {
      "type": "local",
      "command": "npx",
      "tools": ["*"],
      "args": ["-y", "@ndhkaeru/docforge-mcp@latest", "excel"]
    }
  }
}
```

</details>

<details>
<summary>Cursor</summary>

Go to `Cursor Settings` → `MCP` → `Add new MCP Server`.

Use command type with:

```bash
npx -y @ndhkaeru/docforge-mcp@latest excel
```

Or use the standard config above.

</details>

<details>
<summary>Factory</summary>

Use the Factory CLI:

```bash
droid mcp add docforge-excel "npx -y @ndhkaeru/docforge-mcp@latest excel"
```

Alternatively, type `/mcp` inside Factory droid to open the MCP management UI.

</details>

<details>
<summary>Gemini CLI</summary>

Use Gemini CLI MCP configuration with the standard config above, or add a server with:

```bash
gemini mcp add docforge-excel -- npx -y @ndhkaeru/docforge-mcp@latest excel
gemini mcp add docforge-md -- npx -y @ndhkaeru/docforge-mcp@latest md
```

</details>

<details>
<summary>Goose</summary>

Go to `Advanced settings` → `Extensions` → `Add custom extension`.

Use type `STDIO` and command:

```bash
npx -y @ndhkaeru/docforge-mcp@latest excel
```

</details>

<details>
<summary>Junie</summary>

Use Junie's MCP flow or add to `.junie/mcp/mcp.json`:

```json
{
  "mcpServers": {
    "DocForge Excel": {
      "command": "npx",
      "args": ["-y", "@ndhkaeru/docforge-mcp@latest", "excel"]
    }
  }
}
```

</details>

<details>
<summary>Kiro</summary>

Follow Kiro MCP Servers documentation. Example `.kiro/settings/mcp.json`:

```json
{
  "mcpServers": {
    "docforge-excel": {
      "command": "npx",
      "args": ["-y", "@ndhkaeru/docforge-mcp@latest", "excel"]
    }
  }
}
```

</details>

<details>
<summary>LM Studio</summary>

Go to `Program` in the right sidebar → `Install` → `Edit mcp.json` and use the standard config above.

</details>

<details>
<summary>opencode</summary>

Example `~/.config/opencode/opencode.json`:

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "docforge-excel": {
      "type": "local",
      "command": ["npx", "-y", "@ndhkaeru/docforge-mcp@latest", "excel"],
      "enabled": true
    }
  }
}
```

</details>

<details>
<summary>Qodo Gen</summary>

Open Qodo Gen chat panel in VS Code or IntelliJ → Connect more tools → Add new MCP → paste the standard config above.

</details>

<details>
<summary>VS Code</summary>

Follow VS Code MCP server documentation and use the standard config above.

You can also use the VS Code CLI:

```bash
code --add-mcp '{"name":"docforge-excel","command":"npx","args":["-y","@ndhkaeru/docforge-mcp@latest","excel"]}'
```

</details>

<details>
<summary>Warp</summary>

Go to `Settings` → `AI` → `Manage MCP Servers` → `+ Add`, then paste the standard config above.

</details>

<details>
<summary>Windsurf</summary>

Follow Windsurf MCP documentation and use the standard config above.

</details>

### Configuration

DocForge's npm launcher accepts exactly one server argument followed by optional server process arguments:

```bash
npx -y @ndhkaeru/docforge-mcp@latest <excel|md|pdf|docx|pptx|csv|html|text|json> [server args...]
```

Direct npm binaries are also available:

```bash
npx -y -p @ndhkaeru/docforge-mcp@latest docforge-excel-tools
npx -y -p @ndhkaeru/docforge-mcp@latest docforge-md-tools
```

| Environment variable | Description |
| --- | --- |
| `DOCFORGE_MCP_BINARY` | Use this binary path when invoking a direct wrapper for one server. |
| `DOCFORGE_EXCEL_TOOLS_BINARY` | Override `excel-tools` binary. |
| `DOCFORGE_MD_TOOLS_BINARY` | Override `md-tools` binary. |
| `DOCFORGE_PDF_TOOLS_BINARY` | Override `pdf-tools` binary. |
| `DOCFORGE_DOCX_TOOLS_BINARY` | Override `docx-tools` binary. |
| `DOCFORGE_PPTX_TOOLS_BINARY` | Override `pptx-tools` binary. |
| `DOCFORGE_CSV_TOOLS_BINARY` | Override `csv-tools` binary. |
| `DOCFORGE_HTML_TOOLS_BINARY` | Override `html-tools` binary. |
| `DOCFORGE_TEXT_TOOLS_BINARY` | Override `text-tools` binary. |
| `DOCFORGE_JSON_TOOLS_BINARY` | Override `json-tools` binary. |

### Native binaries

This package is a JavaScript launcher plus native binaries under `native/<platform>/`:

```text
native/
  win32-x64/
  linux-x64/
  darwin-x64/
  darwin-arm64/
```

Release automation downloads matching GitHub Release assets before publishing this npm package.

### Tools

#### Excel tools

- **convert_to_markdown** - Convert Excel-family files to Markdown without creating a session.
- **excel_load** / **excel_save** / **excel_save_as_copy** - Load, edit, validate, and save OOXML workbooks.
- **excel_read_range**, **excel_get_rows**, **excel_get_cell**, **excel_find_cells** - Targeted workbook inspection.
- **excel_add_sheet**, **excel_delete_sheet**, **excel_rename_sheet**, **excel_copy_sheet**, **excel_move_sheet** - Sheet operations.
- **excel_edit_cells**, **excel_insert_rows**, **excel_delete_rows**, **excel_insert_column**, **excel_delete_column** - Row, column, and cell edits.
- **excel_set_style**, **excel_set_borders**, **excel_set_shape_style**, **excel_update_shape_text** - Cell and DrawingML shape styling.

#### Markdown tools

- **markdown_outline**, **read_markdown_section**, **md_search** - Structure-aware reads.
- **md_insert_section**, **replace_markdown_section**, **md_delete_section**, **md_move_section** - Section edits.
- **md_read_table**, **md_edit_table**, **md_insert_table**, **md_format_table** - Markdown table operations.
- **md_list_links**, **md_validate_links**, **md_update_toc**, **md_frontmatter** - Link, TOC, and metadata helpers.

#### Conversion tools

- **pdf-tools** - PDF text extraction to Markdown.
- **docx-tools** - DOCX to Markdown.
- **pptx-tools** - PPTX slides to Markdown.
- **csv-tools** - CSV to Markdown table.
- **html-tools** - HTML to Markdown.
- **text-tools** - Plain text/Markdown passthrough.
- **json-tools** - JSON/JSONL passthrough.

## Security and limits

DocForge MCP is **not** a security boundary. It runs local file tools with the permissions granted by your MCP client and operating system.

- Create one MCP entry per server to avoid exposing unnecessary tools.
- Prefer targeted read options such as ranges, previews, and max row/column limits.
- Excel edit/save supports OOXML workbooks and validates save flows. Advanced Excel parts such as macros, DrawingML, pivot tables, slicers, VML, and external links are best-effort and may produce validation warnings.
- PDF/DOCX/PPTX conversion is text/Markdown oriented, not pixel-perfect rendering.
- Linux/macOS artifacts are produced by GitHub Actions on native runners; Windows artifacts end in `.exe`.
