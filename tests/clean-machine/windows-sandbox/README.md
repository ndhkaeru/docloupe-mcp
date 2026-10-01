# Clean-machine test (Windows Sandbox)

Checks whether the Excel MCP server binary works on a Windows machine where **nothing is installed**: no Python, no Node, no .NET runtime, and no network. Windows Sandbox provides that machine: a pristine, disposable copy of the host's Windows.

## Files

| File | Runs on | Purpose |
|---|---|---|
| `Run-SandboxTest.ps1` | Host (PowerShell 7 or 5.1) | Stages the binary, the fixture and the test script; writes a `.wsb`; starts the sandbox; waits for and prints the result |
| `Test-McpServer.ps1` | Inside the sandbox (Windows PowerShell 5.1) | Speaks MCP over stdio to the binary and runs the scenario; writes `result.json` after every step |
| `New-TestWorkbook.ps1` | Host | Generates `fixture.xlsx` (rich-text cell with 3 runs, Vietnamese text, a formula) with no third-party tools |

## Scenario `legacy` (Python server, `servers/excel`)

| Step | What it checks |
|---|---|
| 1 | Start the binary and get an `initialize` response (startup time is recorded) |
| 2 | `tools/list` contains the required tools |
| 3 | `excel_load`, then A1 has 3 rich-text runs with exact Vietnamese text |
| 4 | `excel_edit_cells` B2 with Vietnamese text. Tries the A1 form first; any form the binary rejects is reported in the step detail |
| 5 | `excel_save` with `verify_preservation=true` → `completed` and `preservation_ok` |
| 6 | Reload the saved file: B2 equals the written text code point for code point; A1 still has 3 runs |
| 7 | `excel_validate_workbook` on the saved file |

The sandbox runs with networking disabled, so the binary must work offline.

## Usage

```powershell
# One-time: enable Windows Sandbox (elevated PowerShell), then restart Windows.
Enable-WindowsOptionalFeature -Online -FeatureName Containers-DisposableClientVM -All

# Test the local build (default: <repo>\dist\excel-tools.exe)
.\Run-SandboxTest.ps1

# Test the released binary that the npm launcher downloads
.\Run-SandboxTest.ps1 -Binary "$env:LOCALAPPDATA\docloupe-mcp\v1.1.2\win32-x64\excel-tools.exe"

# Keep the sandbox open afterwards to inspect it
.\Run-SandboxTest.ps1 -KeepOpen

# Only check the harness on this machine (NOT a clean-machine result)
.\Run-SandboxTest.ps1 -RunOnHost
```

**Exit codes:**

| Code | Meaning |
|---|---|
| `0` | Passed |
| `1` | Failed, or timed out |
| `2` | Windows Sandbox is not enabled |
| `3` | A sandbox instance is already running |

**Artifacts** are written to `%TEMP%\docloupe-sandbox-<timestamp>\xt-out`:

- `result.json` and `summary.txt`
- `server-stderr.log`
- `saved.xlsx`

`result.json.machine` records what was present: `python`, `py`, `node`, `dotnet`, the .NET directory, and the VC++ redistributable. In a real sandbox run these should all be absent. `python` may resolve to the Microsoft Store alias stub under `WindowsApps`, which is not a Python install.

## Notes

- The scripts are ASCII-only; Vietnamese literals are JSON escapes or XML character references, so they behave the same under PowerShell 5.1 and 7.
- **BOM handling:** on a UTF-8 console, .NET Framework writes a BOM to a child's stdin. `Test-McpServer.ps1` switches to a BOM-less encoding before starting the server.
- **Case-sensitive JSON:** tool payloads are parsed with a case-sensitive parser, because the legacy server emits keys that differ only by case (`vAlign`/`valign`), which `ConvertFrom-Json` in PowerShell 5.1 rejects.
- **For the C# server:** add a scenario to `Test-McpServer.ps1` (new tool names: `excel_open`, `excel_apply`, `excel_save`…) and run it with `-Scenario`. See `servers/excel_cs/docs/06-testing-and-roadmap.md` §2.10.
