# P2a writer-core decision — 2026-10-01

**Decision:** `System.Xml` is the sole mutation authority for the P2a writer. The original part bytes remain the lexical source of truth. Open XML SDK is **not** the writer, ZIP/package loader, or verifier reader. Its bounded role is a separate G2 schema-validation adapter over detached typed roots; no SDK serialization is committed to an output part. Typed SDK construction of write fragments is deferred unless a later test proves that the exact fragment, without a second hand-written edit, passes G3 and the independent outside-edit byte check.

## Evidence

- S2b's SDK `Worksheet` mutation did not produce its output: the edit was re-applied manually to `XmlDocument`, then a serialized cell was spliced into the original part. Thus S2b proved the `System.Xml` path only. S2 showed that SDK whole-sheet serialization rewrote the default worksheet namespace to `x:`; its package API could not open 3/8 prefixed-content-types fixtures.
- The `P2aSdkRole` probe read the first worksheet from all eight original fixtures and their S2b copies directly from ZIP, made detached typed roots without opening an SDK package, and ran `OpenXmlValidator(Microsoft365)` on each. Error counts before/after: 0/0 for seven fixtures and 1/1 for fixture 01; the *error ID multisets* matched in all eight. A deliberately invalid direct child of a detached worksheet produced `Sch_InvalidElementContentExpectingComplex`, except on fixture 01, where a baseline error masks it (see review finding 1). This establishes an available detached schema check, **not** complete G2 coverage or semantic preservation. The probe copies children, root non-namespace attributes, and prefixed namespace declarations; the SDK rejects an empty prefix in `AddNamespaceDeclaration`, so an explicit source default-namespace declaration cannot be reproduced on its typed root. That is harmless for a URI-based schema check but disqualifies it as a lexical writer. Production G2 must compare stable locations and diagnostics rather than IDs alone and prove coverage per touched/added part.
- The independent `Verify` project remains SDK-free. SDK validation cannot replace G1 relationships/content types, G3 namespace and lexical checks, G4 intent readback, G5 preservation, or an Excel-open/repair oracle. Any touched part unsupported by detached validation is an explicit required `unverified` gap, not `verified`; the save policy in 04 §4 applies.

## P2a implementation boundary

- The writer resolves a target cell namespace-aware, cross-checks its `r` and raw span, applies the mutation **once** to its `XmlDocument`, serializes only the changed subtree, and splices it into the original part. Several nonoverlapping changes in one sheet require a planned set of spans and descending-offset replacement; ambiguous, overlapping, missing or encoding-unsupported spans fail closed. A separate `Verify` reader must check bytes outside the declared spans without reusing the writer's locator.
- No `SpreadsheetDocument.Open`, `OpenXmlWriter`, whole-worksheet `OuterXml`, or SDK-backed package save on the output path. For large sheets, a source-preserving streaming writer requires its own evidence (S4); until then, P2a rejects sheets beyond its proven DOM limit instead of falling back to a lossy writer.
- This is an architecture decision, **not P2a completion**: arbitrary cells, multiple edits, cell-type conversion, staging/G1–G5, and the mutation/corruption matrices remain to be implemented and tested.

## Review findings (2026-10-01)

The original probe passed as soon as **one** fixture detected the injected child (`regressionDetected` was set by the first hit). Re-run per fixture, the injected child was **not** detected on `01-audit-87-source.xlsx`. The probe now injects an error at three depths into every fixture and fails unless each outcome matches the rule below. Result: 24/24 as predicted (`evidence/probe-results.txt`).

1. **Baseline errors mask new errors in the same parent element.**
   - **What happens:** `OpenXmlValidator` reports only the first content-model error per parent element. Fixture 01 already has one at `/x:worksheet[1]` (an unexpected child), so a second misplaced child of `<worksheet>` is never reported. Comparing errors by location or ID cannot reveal it, because it is simply absent.
   - **Scope of the masking:** errors under other parents are still reported. In fixture 01, a misplaced child of `<sheetData>` and an `<f>` after the value inside a cell were both detected.
   - **Resulting G2 rule** (04 §3.2): an edit that changes the child list of element E is covered by G2 only if E has no baseline content-model error. Otherwise it is a required `unverified` gap.
2. **Detached validation does not check relationship references.** A hyperlink with `r:id="rId999"` and no such relationship produced no error. Dangling `r:id` references in touched parts must be checked by G1 (04 §3.1).
3. **G2 catches child-order mistakes:** `<v>` before `<f>` inside a cell, and `<sheetData>` before `<sheetViews>`, were both detected. With hand-written System.Xml edits, G2 is therefore the safety net for element order. The engine still needs schema-derived element-order tables to insert children at the right position in the first place; G2 only catches mistakes.
4. **Fixture 01 is schema-invalid at baseline**, like the Excel-refused fixtures 02 and 05 (S2b review finding 9). Record this alongside the local fixture manifest.

Reproduce from `servers/excel_cs` after S2b's probe has created its ignored output files:

```powershell
dotnet run --project spikes/P2aSdkRole/P2aSdkRole.csproj -- 'D:\data-test\excel-preservation-fixtures\sources' 'spikes/S2b/out'
```
