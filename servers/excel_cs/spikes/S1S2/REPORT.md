# S1 + S2 spike — 2026-10-01

**Scope:** Windows `win-x64`, .NET SDK 10.0.302, Open XML SDK 3.5.1, official MCP SDK 1.4.1.

This is a disposable experiment, **not** the port and not a preservation verifier.

- Source workbooks under `D:\data-test` were only read. The seven sources listed in the external manifest still match their recorded SHA-256 values (re-checked after the review).
- Generated workbooks stay ignored under `out/`. Under decision Q3 they are derived from external fixtures and must not be committed.
- Text evidence that contains no fixture content is committed under `evidence/`.

## S2 — decision: B, conditional on S2b

### Method

`Program.cs probe` runs two variants per fixture: a no-op forced worksheet serialization, and an edit of the first existing cell. Each variant uses two writers:

- **A:** copy the file, open `SpreadsheetDocument` editable, save the worksheet, close the package.
- **B:** read a detached worksheet DOM, write only that ZIP entry, and copy the decompressed bytes of every other entry.
  - For the three sources whose `[Content_Types].xml` uses `<ns0:Types>`, B normalizes the content types **only in an in-memory read view** handed to the SDK; the output copies the original bytes.
  - B still obtains the DOM through `SpreadsheetDocument.Open` on that in-memory view. This is test scaffolding, not a production PackageStore.

`Program.cs probe-autosave` re-runs A with `OpenSettings.AutoSave` on and off (added after the review).

### Results

From the Native AOT executable (`evidence/probe-results.txt`) and `evidence/autosave-check.txt`:

| Source | A (default `AutoSave = true`) | A with `AutoSave = false` | B: untouched entry bytes and `[Content_Types].xml` | B: edited cell / other expanded-name XML |
|---|---|---|---|---|
| `00-base.xlsx` | `xl/workbook.xml` rewritten | identical | identical | edit present; rest equivalent |
| `01-audit-87-source.xlsx` | `xl/workbook.xml` rewritten | identical | identical | edit present; rest equivalent |
| `02-table-metadata-source.xlsx` | cannot open: `Required Types tag not found` | cannot open | identical | edit present; rest equivalent |
| `03-print-area-source.xlsx` | `xl/workbook.xml` rewritten | identical | identical | edit present; rest equivalent |
| `04-rich-text-phonetic-source.xlsx` | `xl/workbook.xml` rewritten | identical | identical | edit present; rest equivalent |
| `05-advanced-package-source.xlsm` | cannot open: `Required Types tag not found` | cannot open | identical | edit present; rest equivalent |
| `06-book1-richtext-source.xlsm` | `xl/workbook.xml` rewritten | identical | identical | edit present; rest equivalent |
| `07-real-package-source.xlsx` | cannot open: `Required Types tag not found` | cannot open | identical | edit present; rest equivalent |

### Findings

1. **The `xl/workbook.xml` rewrite under A is an `AutoSave` artifact, not a property of A.**
   - `SpreadsheetDocument.Open(path, true)` defaults to `AutoSave = true`.
   - `FirstPart` must load the workbook DOM to find the first sheet, and that loaded part is saved again on dispose.
   - With `AutoSave = false`, A left every untouched entry byte-identical on all five fixtures it can open.
   - Untouched-part stability therefore does **not** distinguish A from B. An earlier version of this report claimed otherwise.
2. **A cannot open 3/8 fixtures.** Their `[Content_Types].xml` uses a namespace prefix (`<ns0:Types>`), which the `System.IO.Packaging` loader rejects. These files appear to come from the legacy Python writer (ElementTree). Whether Excel opens them is untested; this is part of S2b. Either way, the production reader must handle them without rewriting the original.
3. **Touched-part lexical form changes.**
   - Open XML SDK writes known namespaces with its own prefix. In all eight originals, the touched worksheet's default `xmlns="…/main"` became `xmlns:x="…/main"` with `x:`-prefixed elements. Under A with `AutoSave`, `xl/workbook.xml` was rewritten the same way and gained an XML declaration.
   - `OuterXml` (B) writes no XML declaration. The 00-base sheet had none, but Excel-authored parts usually carry `<?xml … standalone="yes"?>`.
   - Expanded names and values were preserved (`other_xml=same`).
   - Extension namespaces kept their prefixes: in the synthetic case, `<x:worksheet>`, `<x:c>`, `mc:Ignorable="x14ac z"`, `x14ac:dyDescent` and an unknown `extLst/ext/z:opaque` all survived in both modes.
4. **The SDK does not reject EX-04.** In `EX04_invalid.xlsx`, `mc:Ignorable` names an undeclared `missing` prefix, and neither mode rejects or repairs it (the probe reports `mc=INVALID`). An independent G3 must reject such output. Both synthetic files are incident *shapes* derived from `00-base.xlsx`, not the historical failing binaries.
5. **Limits of `Compare()`:**
   - `other_xml` is a structural comparison of expanded element/attribute names and values. It ignores namespace declarations and whitespace-only text nodes (`XDocument.Parse` defaults). It is **not** XML C14N.
   - The edited cell is excluded by its `r` attribute. Cells without `r`, which OOXML allows, would all be excluded; none of the fixtures has them.
   - The edit check verifies the cell text only, not its `s` or other attributes.
   - ZIP container bytes (compression, metadata) were not compared; only each entry's decompressed bytes.

### Decision

**B**, because:

- A cannot open 3/8 fixtures;
- A gives no control over content types and relationships (01 §5.1).

This is an architectural direction, **not a pass of all S2 exit criteria**.

Under the G3 lexical-fidelity rule decided on 2026-10-01 (04 §3.3), touched parts must keep the source's prefixes, root declarations and XML declaration. Plain SDK serialization does not do that. B is not trusted for P2 until **S2b** shows the following:

- part DOMs built from raw XML without `System.IO.Packaging`;
- a prefix-restoration pass that makes touched parts match O lexically, except for the edited elements;
- `ns0:` content types read without being rewritten;
- an independent G3 that rejects EX-04.

## S1 — Windows host result

`dotnet publish` with Native AOT `win-x64` succeeded with **zero trim/AOT warnings**. The source was re-published after the `probe-autosave` change, and still produced zero warnings (`evidence/publish-recheck.txt`, 41,728,512-byte executable).

Build notes:

- The first attempt failed because `vswhere.exe` was not on `PATH`; adding the Visual Studio Installer directory fixed linker discovery.
- `Host` logging to stdout was disabled.
- A source-generated JSON context was required for the MCP `CallToolResult`, and the tool explicitly emits `structuredContent`.

The binary completed the MCP `initialize` and `tools/call` handshake and returned `{ "Name": "Scores", "FirstCell": "Name" }` on `00-base.xlsx`.

| Metric | Observed |
|---|---:|
| `S1S2.exe` (single-file AOT) | 41,724,928 bytes (39.79 MiB); legacy PyInstaller `excel-tools.exe` is ~25.5 MB |
| Optional debug-symbol `.pdb` (not needed at runtime) | 184,373,248 bytes |
| Ten fresh processes, start → `initialize` reply | min 84.9 ms; median 87.8 ms; max 634.8 ms (first run) |
| Ten fresh processes, start → read-tool reply | median 97.7 ms |

- **What the timings are:** local sequential runs (`evidence/smoke.txt`; the median is the lower median of 10), not cold-machine boot measurements.
- **Dependencies:** `dumpbin /dependents` (`evidence/dumpbin.txt`) lists `ADVAPI32`, `bcrypt`, `CRYPT32`, `IPHLPAPI`, `KERNEL32`, `ncrypt`, `ole32`, `WS2_32` and seven `api-ms-win-crt-*` API-set DLLs (`heap`, `math`, `string`, `convert`, `stdio`, `runtime`, `locale`). No managed .NET runtime DLL appears in the direct imports.
- **Still open:** this is **not** the clean-machine (S9) suite. The binary has not been run in Windows Sandbox yet; the harness in `tests/clean-machine/windows-sandbox` needs a scenario for `read_first_sheet`. macOS and Linux AOT builds cannot be produced on this Windows host. The four-RID S1 exit criterion and S9 remain open, and both are prerequisites for switching the launcher (P7).
- **Size:** has not been tuned (e.g. `OptimizationPreference=Size`).

## Reproduce

From `servers/excel_cs` in PowerShell, with the .NET 10 SDK and Visual Studio C++ Build Tools installed:

```powershell
$env:PATH = 'C:\Program Files (x86)\Microsoft Visual Studio\Installer;' + $env:PATH
dotnet publish spikes/S1S2/S1S2.csproj -c Release -r win-x64 --self-contained true -o spikes/S1S2/out/aot -p:PublishAot=true -p:StripSymbols=true
& spikes/S1S2/out/aot/S1S2.exe probe 'D:\data-test\excel-preservation-fixtures\sources' 'spikes/S1S2/out/aot-probe'
& spikes/S1S2/out/aot/S1S2.exe probe-autosave 'D:\data-test\excel-preservation-fixtures\sources' 'spikes/S1S2/out/autosave-check'
& spikes/S1S2/smoke.ps1 -Runs 10
& 'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Tools\MSVC\14.44.35207\bin\Hostx64\x64\dumpbin.exe' /dependents spikes/S1S2/out/aot/S1S2.exe
```

Committed evidence lives in `evidence/`: `probe-results.txt`, `autosave-check.txt`, `smoke.txt`, `dumpbin.txt`, `publish-final.txt`, `publish-recheck.txt`. Generated workbooks stay untracked under `out/`.
