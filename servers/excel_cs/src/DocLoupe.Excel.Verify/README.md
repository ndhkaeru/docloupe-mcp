# P1 read/verify foundation (incomplete)

This library does not reference Open XML SDK or `System.IO.Packaging`. `WorkbookReader.Peek` follows the OPC root `officeDocument` relationship to the workbook, resolves case-insensitive and percent-encoded part names (including nested workbooks), and reads a bounded raw-cell preview directly from ZIP entries. `VerifyPartial` reports missing/invalid root relationships or workbook parts but does not claim full relationship validation. `MarkupCompatibilityVerifier` checks namespace and markup-compatibility tokens using `XmlReader` independently of the S2b writer.

`VerifyPartial` intentionally returns `unverified` when its current G1 subset and G3 pass. It is **not** a full package validity, schema, intent, preservation or advanced-part verifier. Never present its result as a complete verified workbook. See `../../spikes/S2b/REPORT.md` for evidence and the remaining P1 gates.
