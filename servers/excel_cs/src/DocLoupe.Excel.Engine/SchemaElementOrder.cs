namespace DocLoupe.Excel.Engine;

public static class SchemaElementOrder
{
    public static readonly string[] Worksheet = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
    public static readonly string[] Workbook = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];
    public static readonly string[] SheetData = ["row"];
    public static readonly string[] Row = ["c", "extLst"];
    public static readonly string[] Cell = ["f", "v", "is", "extLst"];
}
