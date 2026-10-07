# Bounded cell-edit parity (P2a/P2b)

Run from the repository root on a machine with .NET 10 and `uv` (Windows or Linux):

```powershell
uv run --no-project --with 'mcp==1.28.1' --with 'openpyxl==3.1.5' --with 'lxml==6.1.1' python servers/excel_cs/tools/parity.py
```

The runner builds the C# server and six synthetic workbooks, then generates a seventh workbook independently of the C# fixture generator with pinned `openpyxl` in a temporary directory. The seventh source has rich text in untouched A1, formatting on edited B1, and formatting on value-cleared D3. It invokes the **actual** C# MCP stdio tools (`excel_open`, `excel_apply`, `excel_save` copy, `excel_close`) and the existing Python implementation (`excel_load`, `excel_edit_cells`, `excel_save_as_copy`, `excel_close`). The shared scenario overwrites a number and cached-formula cell, clears an inline string, writes a Boolean and decimal, sets a formula with numeric cache, inserts new cells, broadcasts a constant using `fill`, and writes a 1×2 rectangle with `set_values`. Python's `excel_edit_cells` represents the same final scalar edits; this is **not** operation-contract parity for `fill` or `set_values`.

On the `openpyxl-rich-style` variant, the C# MCP first checks `excel_apply.expect.rich` on A1 in dry-run without changing the source; its `excel_save` call also requires `equals.value` and `equals.rich` for A1 on the staging file. `parity.py` then opens each output ZIP independently, resolves workbook, sheet, shared-string and style parts via OPC relationships, and checks all edited cell values/formulas and every other cell's value/formula against the source. On the `openpyxl-rich-style` source it also compares the untouched A1 rich-text runs, B1's font/fill/border and cell format after editing its value, and D3's formatting after clearing its value. The check treats `<b/>` and `<b val="1"/>` as equivalent. Negative controls remove one rich run's formatting and reset B1's style, and must both be detected. It does not trust either server's readback or save report. It is a bounded semantic comparison, **not** full package preservation, parity of rich/style mutation operations, formula recalculation, or an oracle of Excel's rendering. The source remains untouched; `verify_preservation=False` is used only for the legacy copy-save because its independent verification is not equivalent to C# G1–G7.

Six synthetic variants were observed locally on Windows and Docker Ubuntu 24.04 with .NET 10 / Python 3.12 on 2026-10-02. The Linux run installs `python3-venv` and the three pinned Python dependencies into an ephemeral container; it does not modify the working tree. The independently generated `openpyxl-rich-style` variant and both negative controls were observed locally on Windows and Docker Ubuntu 24.04 on 2026-10-07; its additional MCP G7 `equals.rich` assertion passed on both systems on 2026-10-08. macOS and remote CI have not yet been run for these additions:

| Synthetic source | C# output | Legacy Python output |
|---|---|---|
| `new-shared-strings` | Pass | Pass |
| `default`, `bom-crlf-standalone` | Pass | The **unmodified** A1 changes from `hello` to `hellohe` (phonetic text folded into the shared-string value) |
| `prefixed-x`, `nested-workbook` | Pass | The namespace-aware OOXML reader no longer finds the **unmodified** A1; cleared D3 is absent rather than an empty cell |

| `opc-percent-case` | Pass | Fails to resolve `Sheet1` from the original case/percent-encoded package |
| `openpyxl-rich-style` (Windows/Docker Linux) | Pass, including rich/style preservation | Pass, including rich/style preservation |

Known divergences are asserted **by exact fixture and observed cell/error**, not silently skipped. If Python changes behavior, the scenario fails so its exception list must be reviewed. These cases must not be counted as Python/C# equivalence. The old prefixed-cell V-06 silent-edit-loss scenario has a separate regression test; this runner does not prove or disprove that case. Add rich/style mutation operations, broader facet coverage, and independently validated fixture families before closing P1/P2b parity.
