# Bounded cell-edit parity (P2a/P2b)

Run from the repository root on a machine with .NET 10 and `uv`:

```powershell
uv run --no-project --with 'mcp==1.28.1' --with 'openpyxl==3.1.5' --with 'lxml==6.1.1' python servers/excel_cs/tools/parity.py
```

The runner builds the C# server and six synthetic workbooks in a temporary directory. It invokes the **actual** C# MCP stdio tools (`excel_open`, `excel_apply`, `excel_save` copy, `excel_close`) and the existing Python implementation (`excel_load`, `excel_edit_cells`, `excel_save_as_copy`, `excel_close`). The shared scenario overwrites a number and cached-formula cell, converts an inline string, writes a Boolean and decimal, inserts new cells, broadcasts a constant using `fill`, and writes a 1×2 rectangle with `set_values`. Python's `excel_edit_cells` represents the same final scalar edits; this is **not** operation-contract parity for `fill` or `set_values`.

`parity.py` opens each output ZIP independently, resolves workbook, sheet, and shared-string parts via OPC relationships, and checks all edited cell values/formulas and every other cell's value/formula against the source. It does not trust either server's readback or save report. It is a bounded semantic comparison, **not** full package preservation, formula recalculation, rich/style parity, or an oracle of Excel's rendering. The source remains untouched; `verify_preservation=False` is used only for the legacy copy-save because its independent verification is not equivalent to C# G1–G7.

Observed locally on Windows, 2026-10-02:

| Synthetic source | C# output | Legacy Python output |
|---|---|---|
| `new-shared-strings` | Pass | Pass |
| `default`, `bom-crlf-standalone` | Pass | The **unmodified** A1 changes from `hello` to `hellohe` (phonetic text folded into the shared-string value) |
| `prefixed-x`, `nested-workbook` | Pass | The namespace-aware OOXML reader no longer finds the **unmodified** A1 in the output |

| `opc-percent-case` | Pass | Fails to resolve `Sheet1` from the original case/percent-encoded package |

Known divergences are asserted **by exact fixture and observed cell/error**, not silently skipped. If Python changes behavior, the scenario fails so its exception list must be reviewed. These cases must not be counted as Python/C# equivalence. The old prefixed-cell V-06 silent-edit-loss scenario has a separate regression test; this runner does not prove or disprove that case. Add more operations, style/rich checks, and separate independently validated fixture families before closing P1/P2b parity.
