# Synthetic P2a fixtures

The .NET generator in this directory creates deterministic-source OOXML workbooks at a caller-supplied directory. CI generates them under temporary test directories; it never reads `D:\data-test`.

Run from the repository root: `dotnet run --project servers/excel_cs/fixtures/DocLoupe.Excel.Fixtures.csproj -- <output-directory>`.

| Variant | Covered feature |
|---|---|
| `default.xlsx` | Default SpreadsheetML namespace; shared string with phonetic run; cached formula; inline string; x14ac and mc prefixes |
| `prefixed-x.xlsx` | `x:` SpreadsheetML and `ns0:Types` content types |
| `bom-crlf-standalone.xlsx` | UTF-8 BOM, CRLF and `standalone="yes"` declaration |
| `opc-percent-case.xlsx` | Case-insensitive part names and percent-encoded sheet relationship target |
| `new-shared-strings.xlsx` | No original SST; first text write must add the part, relationship and content-type override |
| `nested-workbook.xlsx` | Workbook outside the default location; new shared-strings part and relationship remain relative to that workbook |

`DocLoupe.Excel.Engine.Tests` builds these sources in CI and exercises mutation, readback, package integrity, independent byte preservation and detached schema validation. The source generator and test cases are repository-local; no Excel installation or private data is needed. The byte-identical ZIP archive across different OS runtimes is not assumed.
