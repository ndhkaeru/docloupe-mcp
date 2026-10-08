# P2a local acceptance — 2026-10-08

Scope: only the thin `set_value` path defined in [06 §4](06-testing-and-roadmap.md#4-roadmap): open, bounded cells read, apply, staging save through G1–G5, independent readback, and close. This report does **not** close P1, P2b, P6, or P7 or authorize changing the Python launcher.

| Check | Result |
|---|---|
| G1–G5 fixture matrix: six synthetic plus eight local sources | Windows: **14 passed, 0 gaps, 0 failed** for the matrix's exact cell edits. The local sources are outside Git. |
| G2 baseline masking on local fixture 01 | The original worksheet still has an `Sch_UnexpectedElementContentExpectingComplex` error at `/x:worksheet[1]`. The known root content-model error cannot hide a new root child for these edits: ordered direct-child names, namespaces, and attributes and root attributes are unchanged. Corruption tests still report a gap for a new root child or altered direct-child attributes, detect new descendant errors, and retain the gap on an edited cell with its own baseline error. |
| Staging save for local fixture 01 | `ExcelSessions.Open` → `Read` → `Apply` (three `set_value` edits) → `Save(copy)` returned `verified`; the source hash stayed unchanged. This does not imply the original schema-invalid workbook is repaired or can be opened by Excel. |
| Full solution tests | Windows with local corpus: **715 passed**. Docker Linux without the local corpus: **698 passed, 9 skipped**. |
| Independent commands | Schema-order and MCP SDK stdio smoke passed on Windows and Docker Linux; the Linux synthetic-only matrix passed **6/6**. |

From the repository root, with .NET 10 installed, reproduce the private-corpus check without adding its source files to Git:

```powershell
$env:DOCLOUPE_P2A_LOCAL_FIXTURES = 'D:\data-test\excel-preservation-fixtures\sources'
dotnet test servers/excel_cs/DocLoupe.Excel.slnx -c Release
dotnet run --project servers/excel_cs/tools/FixtureMatrix/FixtureMatrix.csproj -c Release -- "$env:TEMP\excel-cs-p2a-matrix.json" $env:DOCLOUPE_P2A_LOCAL_FIXTURES
dotnet run --project servers/excel_cs/tools/SchemaOrder/SchemaOrder.csproj -c Release
dotnet run --project servers/excel_cs/spikes/S3/S3.csproj -c Release -- servers/excel_cs/src/DocLoupe.Excel.Server/bin/Release/net10.0/DocLoupe.Excel.Server.dll
```

**Outstanding cross-platform evidence:** `.github/workflows/excel-cs.yml` declares Windows x64, Linux x64, macOS x64, and macOS arm64, but this unpushed branch has not run those four remote runners. No local macOS run was performed. The private corpus is not in CI; its acceptance evidence is local Windows only. Do not label P2a cross-platform complete until the remote matrix and MCP smoke are green on all four runners. Excel-open/repair testing of baseline-invalid fixtures is a separate oracle/release concern and is not inferred from G1–G5.
