# 01 — Architecture and stack

## 1. Goals and non-goals

**Goals**

1. Edit OOXML workbooks (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) **without changing anything that was not requested**.
2. Every save **proves** three things: (a) the intent reached the file, (b) everything else is unchanged, and (c) the package is valid. Anything that cannot be proven is reported as unverified.
3. Be easy for agents to drive: few tools, strict schemas, A1 addresses, and results expressed the same way the agent expressed the request.
4. Handle large files (millions of cells, files of several hundred MB) within bounded memory.
5. Replace the Python server without changing the npm launcher (same binary name and asset name).

**Non-goals (for now)**

- Editing `.xls` / `.xlsb`. Reading them only for Markdown conversion is optional (see §11).
- Evaluating formulas inside the server. Recalculation is delegated to an external oracle (Excel or LibreOffice). Without an oracle, results are reported as `unverified`.
- Creating pivot tables, slicers, timelines, or threaded comments. These are **preserved and verified**, but there are no ops that create them.
- Running macros. VBA is never executed, not even by an oracle.

## 2. Stack

| Component | Choice | Role | Notes / to verify |
|---|---|---|---|
| Runtime | .NET 10 (LTS), C# 14 | Whole server | — |
| MCP | `ModelContextProtocol` (official C# SDK), stdio transport | Tool registration, JSON-RPC, cancellation, `structuredContent` | Spike **S3**: `UseStructuredContent` API, `CancellationToken` plumbing, Native AOT with explicit tool registration |
| OOXML DOM | `DocumentFormat.OpenXml` 3.x (Open XML SDK) | Typed DOM per part: worksheet, styles, sharedStrings, table, drawing… | Spike **S1** (AOT/trimming) and **S2** (prefixes, `mc:Ignorable`, and unknown elements preserved) |
| Package (ZIP) | `System.IO.Compression.ZipArchive` plus a custom `PackageStore` | Per-entry ZIP read/write; relationship and content-type model | `System.IO.Packaging` is not used for writing because it regenerates `[Content_Types].xml` (see §5.1) |
| Independent reader for the verifier | `XmlReader` + `ZipArchive` | Re-reads the written file without going through Open XML SDK | **Deliberately shares no code** with the writer |
| JSON | `System.Text.Json` with the source generator | Input/output DTOs | Required for AOT |
| Number format → display string | `ExcelNumberFormat` (MIT) | Produces `display` from a value and a numFmt code | Cross-checked against Excel with fixtures (S7) |
| Formulas | In-house tokenizer following the Excel formula grammar; XLParser as a reference | Reference shifting on insert/delete, dependency lookup | Check the license if a third-party parser is used |
| Unicode | `System.Globalization.StringInfo` (UAX #29 graphemes), `string.Normalize` | Rich-text matching and offsets | .NET 5+ uses extended grapheme clusters. Must work with `InvariantGlobalization=true` (§12); spike **S9** |
| Excel oracle | Helper **PowerShell 5.1 script** (built into Windows), embedded in the binary as a resource and driving Excel over COM | Rendering (PNG produced by Excel itself), recalculation, "does Excel ask to repair" | Windows with Office only. Native AOT does not support built-in COM interop, so this runs out of process. Not a .NET helper, because that would need an installed runtime. Spike **S5** |
| LibreOffice oracle | `soffice --headless`, used only if already installed | Rendering (PNG produced by LibreOffice from a temp copy whose print area is the target range), recalculation | All three OSes. Spike **S6** |
| PDF rasterization | **None in the server.** Images are produced by the oracle applications themselves | — | A rasterizer such as PDFium would add native libraries per platform and break the single-file distribution (§12) |
| Testing | xUnit, Verify (snapshots), BenchmarkDotNet, FsCheck (property-based) | — | — |

## 3. Project layout

```
servers/excel_cs/
  README.md
  docs/                               # these documents
  DocLoupe.Excel.sln
  src/
    DocLoupe.Excel.Server/            # MCP host: tool declarations, DTOs, JSON source-gen, server instructions
    DocLoupe.Excel.Model/             # A1 addresses, values, effective style, rich-text markup, semantic paths, hashes
    DocLoupe.Excel.Package/           # PackageStore: ZIP, relationships, content types, byte-preserving writer
    DocLoupe.Excel.Engine/            # Sessions, ledger, planner, ops, DOM and streaming editors
    DocLoupe.Excel.Formula/           # Tokenizer, reference shifting, dependencies
    DocLoupe.Excel.Verify/            # Independent reader, snapshots, comparer, normalization rules, gates
    DocLoupe.Excel.Oracles/           # Oracle interfaces and clients (Excel, LibreOffice)
  oracle/
    excel-oracle.ps1                  # Windows COM helper (PowerShell 5.1), embedded into the binary as a resource
  tests/
    DocLoupe.Excel.Model.Tests/
    DocLoupe.Excel.Engine.Tests/
    DocLoupe.Excel.Verify.Tests/      # Includes the corruption corpus (verifier sensitivity tests)
    DocLoupe.Excel.Parity/            # Same scenarios against the Python and C# servers
    DocLoupe.Excel.Bench/
    fixtures/                         # Fixtures committed to the repo, with manifest.json (sha256)
  tools/
    evidence/                         # CLI that replaces excel_build_preservation_summary
```

**Dependency rule:** `Verify` depends only on `Model` and `Package` (to open the ZIP). It **must not** depend on `Engine` or on Open XML SDK. An architecture test (NetArchTest) enforces this.

## 4. Session model

### 4.1 Components

```
Session
  id                : "xs_" + 16 base32 chars
  source            : absolute, normalized path
  source_fingerprint: { sha256, size, mtime }        # captured at open
  format            : xlsx | xlsm | xltx | xltm
  store             : PackageStore                    # original ZIP entries (read-only) + overlay of edited parts
  doms              : Map<partName, PartDom>          # DOMs are loaded only for parts that have been touched
  ledger            : List<LedgerEntry>               # ordered by revision
  revision          : int                             # 0 = just opened
  baseline_snapshot : lazy; built from O by the independent reader
  lock              : SemaphoreSlim(1)
```

**The core difference from the Python server:** the old server **turns the whole workbook into a dict** (`serialize_excel`), ships it across a worker process as JSON, `deepcopy`s it on edits and checkpoints, and **rebuilds** the file on save. The new server **treats the package itself as the model**. A part that is not edited is never re-parsed and never re-serialized.

### 4.2 Sheet access strategy

| Mode | When | How |
|---|---|---|
| DOM | Sheet part ≤ `DOCLOUPE_EXCEL_DOM_MAX_BYTES` (default 32 MB uncompressed) | Load `Worksheet` with Open XML SDK and edit the DOM |
| Streaming | Larger than the threshold | `OpenXmlReader` → apply the change list keyed by row/cell → `OpenXmlWriter`. Memory scales with the number of changes, not the number of cells |

Both modes implement `ISheetEditor`. A parity test requires both modes to produce semantically identical results for the same ops.

### 4.3 Ledger, revisions and undo

Each successful `excel_apply` appends one `LedgerEntry`:

```
LedgerEntry {
  revision, timestamp,
  ops        : the ops as the agent sent them (normalized),
  resolved   : ops with resolved addresses (absolute coordinates after transforms),
  effects    : List<Effect>      # path, before, after_expected
  transforms : List<Transform>   # row/column insert/delete, sheet rename…
}
```

- **Undo to revision N:** rebuild the overlay by replaying the ledger from O up to N. All ops are deterministic, so replay gives a unique result. For large files, overlay checkpoints can be stored every K revisions.
- **Persistence:** the ledger and fingerprint are written to `<session_dir>/<id>/ledger.json` after every revision. After a server restart, `excel_open` with `resume` replays the ledger if the source fingerprint still matches. Retention is 2 days, the same as the legacy `_SESSION_CHECKPOINT_RETENTION_SECONDS`.

### 4.4 Addressing within one batch

The default is `address_mode: "base"`: **every address in one `excel_apply` call refers to the `base_revision` state**, which is exactly what the agent just read. The engine pushes the addresses of later ops through the transforms of earlier ops.

- Example: insert 2 rows before row 5, then edit `B10`. `B10` means the `B10` the agent read. The engine rewrites it to `B12`, and the readback reports `B10 → B12`.
- If a later op targets something an earlier op deleted, the call fails with `TARGET_DELETED_IN_BATCH`.
- `address_mode: "sequential"` (for experienced callers): each op sees the state left by the previous ops.

This replaces the legacy rule ("clone everything first, then insert bottom-to-top") that agents had to remember themselves.

## 5. Pipelines

### 5.1 PackageStore and writing

- **Open:** read the ZIP entry list, and parse `[Content_Types].xml` and every `.rels` into data. Content parts are **not** parsed yet.
- **Overlay:** an edited part is held as a DOM or as new bytes. Added and deleted parts are recorded as package effects.
- **Writing, entry by entry:**
  - Untouched part: **copy the decompressed content verbatim** from the source, keeping entry order and entry names.
  - Edited part: preserve the original XML outside the declared edit region while serializing the edited subtree from the detached DOM; an XML-aware lexical splice demonstrated this for one cell in S2b. G3 checks namespace/prefix fidelity independently. A whole-part writer for other ops still needs an equivalent guarantee or a separately tested prefix-restoration pass.
    - S2 found that Open XML SDK writes known namespaces with its own prefix: a default worksheet namespace becomes `x:` even when nothing is edited, and `OuterXml` drops the XML declaration.
    - It did keep `mc:Ignorable`, `x14ac:dyDescent` and an unknown `extLst` in the synthetic prefixed case.
    - A plain `OuterXml` write is therefore not sufficient (spike S2b).
  - Part DOMs are built from the raw part XML, **without** `System.IO.Packaging`, which cannot open some valid-looking packages (see the S2 decision below).
  - `[Content_Types].xml` and `.rels`: **edit only the affected entries**. Never regenerate them. This is why `System.IO.Packaging` is not used for writing: it regenerates content types from the part list.
- ZIP-level differences (compression level, timestamps) carry no meaning. They are covered by normalization rule `N-ZIP` (see 04 §5).

**S2 decision (2026-10-01): approach B, conditional on S2b.**

- **Why not A:** approach A (`SpreadsheetDocument` editable) could not open three of the eight external fixtures, whose `[Content_Types].xml` uses a namespace prefix (`<ns0:Types>`). It also gives no control over content types and relationships.
- **Correction:** the `xl/workbook.xml` rewrite first reported for A was caused by `AutoSave` (on by default) combined with reading the workbook DOM. With `AutoSave = false`, A left every untouched entry byte-identical on the five fixtures it could open. Untouched-part stability is therefore not a differentiator.
- **B after S2b:** a direct ZIP/XML DOM prototype parsed all eight fixtures without `System.IO.Packaging`, preserved every byte outside one edited cell on the touched sheet and every decompressed byte of untouched parts, and rejected synthetic EX-04 using an independent G3. This is a single-cell proof, not the general writer.
- **Production PackageStore must:**
  - read prefixed content types without rewriting the originals;
  - build part DOMs without `System.IO.Packaging`;
  - restore source prefixes and declarations;
  - enforce G3/G4/G5.

Evidence and reproduction: `../spikes/S1S2/REPORT.md`, `../spikes/S2b/REPORT.md`. Excel COM opened original and S2b-edited `07` read-only, but refused original and edited `02`/`05`; normalizing only their content-types prefix did not resolve the refusals. Do not treat an XML/package validator pass as proof that Excel opens a workbook.

### 5.2 `excel_apply` pipeline

```
input DTO ──► 1. Schema validation (discriminated union on "op")
          ──► 2. base_revision == session.revision?   (no → REVISION_CONFLICT)
          ──► 3. Resolve targets: A1 / defined name / table / match → absolute coordinates
                 (ambiguous → TARGET_AMBIGUOUS; missing → TARGET_NOT_FOUND)
          ──► 4. Check `expect` preconditions against the base state   (fail → PRECONDITION_FAILED)
          ──► 5. Plan: compute the predicted effects and transforms for the whole batch
          ──► dry_run? ──► return the plan (expected diff); session unchanged
          ──► 6. Apply to copy-on-write copies of the affected parts
          ──► 7. In-session check: re-read every touched facet and compare it with after_expected
                 (mismatch → roll back the batch, INTERNAL_INTENT_MISMATCH: an engine bug, always reported)
          ──► 8. Commit the overlay, append to the ledger, revision++
          ──► output: actual diff + readback + new revision
```

Any failure in steps 1–7 leaves the session **unchanged**. This fixes audit finding EX-09.

### 5.3 `excel_save` pipeline

```
1. Preflight: does the current source fingerprint match the one taken at open?  (no → SOURCE_CHANGED_ON_DISK)
              does the destination exist and differ from the source?            (mode/overwrite policy)
2. Write staging: <destination dir>/.~<name>.<random>.tmp   (same volume, so the rename is atomic)
3. Run the gates on the staging file (details in 04 §3):
     G1 package  G2 schema  G3 markup-compat  G4 intent  G5 preservation
     G6 advanced parts  G7 agent assertions  G8 oracles (optional)
4. A non-overridable gate failed, a G5 failure was not overridden, or a required gap was not accepted (04 §4) → delete staging and return
   SAVE_BLOCKED with the report. The destination does NOT change.
5. Commit: re-check the destination fingerprint → back up the old file (when overwriting)
           → verify the backup hash → File.Move(staging, destination, overwrite)
           → verify hash(destination) == hash(staging)
6. Update the session: source = destination, O = W, ledger marked "saved" at revision R
7. Return the gate report (verified / failed / unverified per gate) + readback of the edited facets
```

Kept from the legacy server, because this part of its design is sound: staging in the same directory, detection of concurrent file changes, backups with hash checks, 2-day backup retention, and blocking when a digital signature would be invalidated.

### 5.4 Independent reader and snapshots

- `Verify.Reader` opens the ZIP with `ZipArchive` and reads each part with `XmlReader` (namespace-aware, DTDs prohibited).
- **The semantic snapshot covers:**
  - cells: value, type, formula and `<f>` attributes, cache, rich runs, phonetics, style resolved down to individual attributes;
  - rows/columns, merges, hyperlinks, comments, validations, conditional formats, tables, defined names;
  - print settings, views, protection, workbook and document properties;
  - drawing anchors, relationships, content types.
- **Nothing is left uncompared.** Whatever the snapshot does not model is compared as **canonical XML**: namespaces compared by URI rather than prefix, attribute order ignored, whitespace normalized except inside elements where whitespace is significant. Parts differ only in how detailed the report is.
- **Prefix fidelity is a separate check.** This semantic comparison ignores prefixes. For parts written by this server, lexical prefix and declaration fidelity is enforced separately by G3 (04 §3.3).
- Snapshots are taken in streaming fashion, without a DOM. For large sheets, the two files are compared as parallel streams, with transforms applied.

## 6. Concurrency, cancellation and timeouts

- **One lock per session.** `excel_read` and `excel_find` also take the lock, but may read from an immutable snapshot of the overlay so they do not block for long.
- **Heavy operations** (open, save, verify, render) go through a global semaphore, `EXCEL_MCP_MAX_HEAVY_WORKERS`.
- **Cancellation:** MCP `notifications/cancelled` becomes a `CancellationToken`. The engine checks it between pipeline steps, and between row batches when streaming. Cancelling during `apply` rolls back; cancelling during `save` deletes the staging file and leaves the destination untouched.
- **Timeouts** are per operation class and keep the legacy names: `EXCEL_MCP_LOAD_TIMEOUT_SECONDS`, `EXCEL_MCP_SAVE_TIMEOUT_SECONDS`, `EXCEL_MCP_VERIFY_TIMEOUT_SECONDS`, `EXCEL_MCP_HEAVY_TIMEOUT_SECONDS`.
- **Oracles** run out of process. On timeout or cancellation the whole process tree is killed (Job Object on Windows, process group on Unix), the same principle as the legacy `excel_capture`.
- **No worker processes needed:** the Python server needs them (`cancellable.py`) because Python threads cannot be cancelled. C# cancels cooperatively in process, so load/save/verify **no longer need workers**.

## 7. Limits and safety

| Item | Default | Setting |
|---|---|---|
| Input size per call | 4 MB | `EXCEL_MCP_MAX_INPUT_BYTES` (legacy name kept) |
| Result size | 256 KB; larger results are paginated or written to a file | `EXCEL_MCP_MAX_RESULT_BYTES` (legacy name kept) |
| Cells returned per `excel_read` | 2,000 (with cursor) | `DOCLOUPE_EXCEL_READ_MAX_CELLS` |
| Total uncompressed package size | 4 GB | `DOCLOUPE_EXCEL_MAX_UNCOMPRESSED_BYTES` |
| Max compression ratio per entry (zip-bomb guard) | 1:200 for entries > 1 MB | `DOCLOUPE_EXCEL_MAX_COMPRESSION_RATIO` |
| Entry count | 20,000 | `DOCLOUPE_EXCEL_MAX_PARTS` |
| XML | `DtdProcessing.Prohibit`, `XmlResolver = null` | — |
| Paths | Normalized to absolute; device paths rejected, UNC rejected unless allowed; optional root allow-list | `DOCLOUPE_EXCEL_ALLOWED_ROOTS` (optional) |
| Macros | Never executed. The Excel oracle opens files with `AutomationSecurity = ForceDisable` | — |
| External links | Never fetched or refreshed | — |

## 8. Observability

- stdout carries **only** JSON-RPC. Logs go to stderr as JSON lines; the level is set with `DOCLOUPE_EXCEL_LOG_LEVEL`.
- Every open/apply/save/verify result carries a `metrics` block: per-phase timings, parts read/written, peak memory. This is the equivalent of the legacy `performance` block.
- Build metadata (`DOCLOUPE_SERVER_VERSION`, `DOCLOUPE_COMMIT_SHA`) is embedded in the assembly and returned by `excel_status`.

## 9. Packaging and release

- **Binary name:** `AssemblyName` is `excel-tools`, so the binary is `excel-tools` / `excel-tools.exe`, the same as the Python build.
- **Build:** `dotnet publish src/DocLoupe.Excel.Server -c Release -r <rid> -p:PublishAot=true` for the RIDs `win-x64`, `linux-x64`, `osx-x64`, `osx-arm64`. These are exactly the four platforms the launcher supports (`platformKey()` in `packages/npm/bin/docloupe-mcp.js`).
- **Self-contained:** users install no runtime or libraries. The remaining OS-level dependencies and the Linux glibc floor are in §12.
- **Runners:** Native AOT cannot cross-compile across operating systems, so each RID is built on a runner of the same OS. The existing matrix (`windows-latest`, `ubuntu-latest`, `macos-15-intel`, `macos-14`) already fits.
- **AOT fallback:** if spike S1 shows that Open XML SDK or the MCP SDK is not yet AOT-ready, fall back to `PublishSingleFile` + `SelfContained` + `ReadyToRun`. This is still a single file, but larger and slower to start.
- **Asset name** stays `docloupe-mcp-excel-tools-<platform>[.exe]`, so the launcher needs no change.
- **Side-by-side period with the Python server:** the launcher only accepts server names listed in `SERVERS`, and file names matching `CACHE_FILE_PATTERN = /^[a-z]+-tools.../`, which forbids hyphens and digits. Use the temporary name **`excelnext`** (binary `excelnext-tools`), add `excelnext` to `SERVERS`, and switch to `excel` at parity (see 06 §4).
- **Binary smoke tests:** must be written fresh because tool names change. The existing `tests/test_excel_binary_smoke.py` applies to the Python build only.

## 10. Configuration summary

| Variable | Meaning | Notes |
|---|---|---|
| `EXCEL_MCP_*_TIMEOUT_SECONDS`, `EXCEL_MCP_MAX_HEAVY_WORKERS`, `EXCEL_MCP_MAX_INPUT_BYTES`, `EXCEL_MCP_MAX_RESULT_BYTES` | As in the legacy server | Names kept so MCP client configs do not change |
| `EXCEL_MCP_MAX_LOAD_RESULT_BYTES`, `EXCEL_MCP_WORKER_PROCESS` | — | Dropped: there are no workers or load artifacts any more |
| `DOCLOUPE_EXCEL_BACKUP_DIR` | Backup directory | Name kept; 2-day retention |
| `DOCLOUPE_EXCEL_SESSION_DIR` | Ledger and checkpoints | New |
| `DOCLOUPE_EXCEL_DOM_MAX_BYTES` | DOM → streaming threshold | New |
| `DOCLOUPE_EXCEL_ORACLE` | `auto` \| `excel` \| `libreoffice` \| `none` | New |
| `DOCLOUPE_SOFFICE_PATH` | LibreOffice path | Replaces the `soffice_path` parameter of `excel_capture` |
| `DOCLOUPE_EXCEL_DISPLAY_LOCALE` | Locale for locale-dependent built-in formats | New; default `en-US` |
| `DOCLOUPE_EXCEL_ALLOWED_ROOTS` | Allowed read/write roots | New, optional |
| `DOCLOUPE_EXCEL_LOG_LEVEL` | Log level | New |

## 11. Formats

| Format | Open for editing | `convert_to_markdown` / `excel_peek` |
|---|---|---|
| `.xlsx`, `.xltx` | ✅ | ✅ |
| `.xlsm`, `.xltm` | ✅ (VBA kept, never run) | ✅ |
| `.xls`, `.xlsb` | ❌ `UNSUPPORTED_FORMAT`, with a conversion hint | Optional later phase: read via ExcelDataReader (MIT). Until it exists, error messages **must not** suggest this path. The audit found the legacy server suggesting a path that does not work |

## 12. Runtime dependencies on user machines

**Requirement:** the binary runs on a machine where **nothing has been installed**: no Node, no .NET runtime, no Python, no VC++ redistributable, no ICU, no LibreOffice. It may only use what the operating system itself ships. Node ≥ 18 is needed **only** for the optional npm launcher, which downloads and caches the binary; it is never needed to run the server (see §12.4). Since the launcher downloads **one file** per platform (`docloupe-mcp-excel-tools-<platform>[.exe]`), the binary must not depend on side-by-side native libraries.

### 12.1 Dependency inventory

| Dependency | Where it would come from | Risk if absent | Decision |
|---|---|---|---|
| .NET runtime | Native AOT compiles it in; the single-file fallback is `SelfContained` | None | Both build modes are self-contained |
| VC++ / C runtime (Windows) | Native AOT uses the Universal CRT that ships with Windows 10+ | None on supported Windows | Minimum OS = the .NET 10 support matrix |
| **ICU (Linux)** | .NET loads `libicu` from the OS for culture data | Startup failure ("Couldn't find a valid ICU package") on minimal distros and containers | **`InvariantGlobalization=true`.** Display formatting uses the server's own locale tables (02 §2.2), which also makes `display` deterministic across machines. Grapheme segmentation (`StringInfo`), NFC normalization and case-insensitive matching must work in invariant mode (spike S9) |
| **OpenSSL (Linux)** | .NET crypto (SHA-256 fingerprints, SHA-512 protection hashes) calls `libssl`/`libcrypto` | Hashing throws on systems without OpenSSL 1.1/3 | **Managed SHA-256/SHA-512 implementation** (verified against NIST test vectors), so the binary has no native crypto dependency |
| zlib (compression) | Statically included in .NET's compression native code | None | Confirm in S9 with `ldd` / `otool -L` |
| **glibc version (Linux)** | Native AOT links against the build machine's glibc | The binary does not start on distros with an older glibc. The current PyInstaller build has the same problem today, because it builds on `ubuntu-latest` without a container | Build the `linux-x64` binary inside an **older-distro container** (e.g. glibc 2.28 class), so it runs on Ubuntu 20.04+/Debian 11+/RHEL 8+. Document the floor |
| musl (Alpine) | — | Not supported: the launcher has no musl platform key | Out of scope; error message from the launcher |
| macOS version | Deployment target of the build | The binary refuses to start on older macOS | Set the minimum explicitly to the .NET 10 floor; document it |
| PDFium / other native libraries | — | Would break single-file distribution | Not used (§2) |
| PowerShell (Excel oracle) | Windows PowerShell 5.1 is part of Windows 10/11 | Oracle unavailable | Run with `-NoProfile -ExecutionPolicy Bypass` (process scope). If Group Policy blocks scripts, report the oracle as `unavailable` rather than failing |
| Microsoft Excel (oracle) | User's Office install | Oracle unavailable | Optional; reported as a gap (04 §3.8) |
| LibreOffice (oracle) | User's install | Oracle unavailable | Optional; never auto-installed. Path from `DOCLOUPE_SOFFICE_PATH` or the standard install locations |
| Fonts (rendering) | OS | LibreOffice substitutes missing fonts (Calibri, Cambria…) | Listed in `fidelity_notes` |

### 12.2 Distribution concerns that are not libraries

| Concern | Note |
|---|---|
| Binary size | Current PyInstaller `excel-tools.exe` is 26.7 MB (25.5 MiB). The S1 spike's Native AOT binary (Open XML SDK + MCP SDK + Hosting, untuned) is 41.7 MB (39.8 MiB). Size tuning (`OptimizationPreference=Size`, dropping Hosting) is still to be tried. It is acceptable either way, because the launcher caches per release |
| Startup | PyInstaller one-file extracts itself to a temp directory on every start. Measured to the `initialize` response (V-12): 2.4 s in a clean Windows Sandbox, but 18–20 s on the development machine, where scanning of the extracted files is the likely cause. It also fails where the temp directory is `noexec`. Native AOT starts in place; target < 150 ms to the first MCP response |
| Code signing | Sign the Windows binary (Authenticode) and the macOS binary (Developer ID). Unsigned executables under `%LOCALAPPDATA%` are blocked by AppLocker/WDAC in some companies and attract antivirus false positives (a known problem with PyInstaller builds) |
| Quarantine / Mark-of-the-Web | The launcher downloads with Node, which does not set the macOS quarantine attribute or MotW, so Gatekeeper and SmartScreen prompts do not apply. Re-check this if the download path changes |

### 12.3 Startup self-check

At startup the server checks its own environment and exposes the result in `excel_status.server.environment`, instead of failing later in the middle of a save:

- the temp and session directories are writable;
- the backup directory is writable;
- oracle availability is probed lazily, on first use, and cached.

A failed hard requirement makes the server exit with one clear stderr line naming the missing item.

### 12.4 Running without Node (direct binary, offline machines)

The supported way to run on a machine with no runtimes is to **copy the binary and point the MCP client at it**:

1. Take the platform asset from the GitHub release, or from an offline bundle (below), and put it anywhere, e.g. `C:\Tools\docloupe\excel-tools.exe` or `~/bin/excel-tools`.
2. Configure the MCP client to start it directly over stdio:

   ```json
   { "mcpServers": { "docloupe-excel": { "command": "C:\\Tools\\docloupe\\excel-tools.exe", "args": [] } } }
   ```

3. Verify the target machine with the built-in CLI before connecting a client:
   - `excel-tools --version` prints the version, commit and RID.
   - `excel-tools --self-check` runs the §12.3 checks plus a full open → apply → save → verify round trip on an embedded fixture, writing only to a temp directory. It exits non-zero with the failing item.

**Offline bundle:** each release also publishes `docloupe-mcp-excel-tools-offline.zip`, containing:
- the four platform binaries;
- `SHA256SUMS`;
- a short `INSTALL.md`.

The machine needs no network access at runtime. The server never downloads anything; only the npm launcher does.

**What still depends on the machine:**

| Item | Without it |
|---|---|
| A supported OS: Windows 10/11 x64, Linux x64 with glibc ≥ the documented floor, macOS ≥ the documented minimum (x64 or arm64) | The binary does not start. Alpine/musl, 32-bit and Windows ARM64 are not supported |
| Excel or LibreOffice | `excel_render`, recalc and `open_check` report `ORACLE_UNAVAILABLE`. Editing, saving and gates G1–G7 work fully, since they need no oracle |
| Execution permission (AppLocker/WDAC, `noexec` home directories) | Blocked by policy. A signed binary and a permitted install path are needed |
