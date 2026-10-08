# P2a fixture/gate matrix

From the repository root:

```powershell
dotnet run --project servers/excel_cs/tools/FixtureMatrix/FixtureMatrix.csproj -c Release -- $env:TEMP\excel-cs-matrix.json D:\data-test\excel-preservation-fixtures\sources
```

Omit the second argument to test only the six committed synthetic fixtures. The JSON report lists G1–G5 separately for each fixture as `passed`, `gap`, or `failed`, including issue details. A failed gate or exception makes the command exit nonzero; a gap remains visible without being reported as a pass. CI publishes synthetic-only JSON separately for each runner. Source packages are not changed and the local corpus is never committed.

This exercises the narrow `set_value` writer and its G1–G5 checks, **not** the full `excel_save` lifecycle, G6/G7, an Excel-open oracle, or all operations. Local fixture 01 has a baseline schema defect. Its modeled cell edits now pass G2 only when the known worksheet-root content-model error cannot mask them (the immediate worksheet children and their attributes are unchanged); a changed root child still reports a gap. The 2026-10-08 local run produced **14 passed, 0 gaps, 0 failed** for the exact edits in this tool. This neither makes the original fixture schema-valid nor guarantees G2 coverage for other edits. Excel rejects the original 02 and 05 files independently of this gate matrix. A signed file can pass G1–G5 yet still be blocked by G6 on save.
