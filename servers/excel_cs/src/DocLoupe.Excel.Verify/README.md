# P1 read/verify foundation (incomplete)

This library does not reference Open XML SDK or `System.IO.Packaging`. `WorkbookReader.Peek` reads workbook relationships and a bounded raw-cell preview directly from ZIP entries. `MarkupCompatibilityVerifier` checks namespace and markup-compatibility tokens using `XmlReader` independently of the S2b writer.

`VerifyPartial` intentionally returns `unverified` when its current G1 subset and G3 pass. It is **not** a full package validity, schema, intent, preservation or advanced-part verifier. Never present its result as a complete verified workbook. See `../../spikes/S2b/REPORT.md` for evidence and the remaining P1 gates.
