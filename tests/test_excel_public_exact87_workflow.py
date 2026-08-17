"""Public-tool creation and preservation coverage for the exact 87 audit keys."""

import json
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers" / "excel"))

import main as M  # noqa: E402


LEGACY_SOURCE = Path(r"D:\data-test\excel-preservation-fixtures\sources\01-audit-87-source.xlsx")
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN = f"{{{SHEET_NS}}}"
REL_ID = f"{{{REL_NS}}}id"


def _load_key(load_result: str) -> str:
    return load_result.split("session_key='")[1].split("'")[0]


def _child(parent: ET.Element | None, name: str) -> ET.Element | None:
    return None if parent is None else parent.find(f"{MAIN}{name}")


def _children(parent: ET.Element | None, name: str) -> list[ET.Element]:
    return [] if parent is None else parent.findall(f"{MAIN}{name}")


def _attrs(node: ET.Element | None) -> dict[str, str]:
    return {} if node is None else dict(node.attrib)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _node_by_attr(parent: ET.Element | None, name: str, attr: str, value: str) -> ET.Element | None:
    return next((node for node in _children(parent, name) if node.get(attr) == value), None)


def _create_exact87_workbook(output_path: Path) -> None:
    created = json.loads(
        M.excel_create_workbook(
            sheet_names=["Scores", "Hidden", "VeryHidden"],
            active_sheet="Scores",
            target_path=str(output_path),
        )
    )
    session_key = created["session_key"]
    try:
        M.excel_edit_cells(
            session_key,
            "Scores",
            [
                {"row_index": 0, "edits": {0: "Name", 1: "Total"}},
                {"row_index": 1, "edits": {0: "Alice"}},
                {"row_index": 2, "edits": {0: "Bob", 1: 20}},
            ],
        )
        M.excel_set_formula(
            session_key,
            "Scores",
            "B2",
            "=SUM(B3)",
            cached_value=10,
            cached_value_present=True,
            cache_policy="replace",
        )
        M.excel_set_auto_filter(
            session_key,
            "Scores",
            ref="A1:B3",
            filter_columns=[{"colId": 0, "filters": ["Alice"]}],
            sort_state={
                "ref": "A2:B3",
                "conditions": [{"ref": "B2:B3", "descending": True}],
            },
            mode="replace",
        )
        M.excel_set_calculation_properties(
            session_key,
            {
                "calcMode": "manual",
                "forceFullCalc": False,
                "fullCalcOnLoad": False,
                "iterate": True,
                "iterateCount": 55,
                "iterateDelta": 0.002,
            },
        )
        M.excel_add_named_style(session_key, "AuditStyle", {"bold": True})
        M.excel_add_defined_name(
            session_key,
            "AuditName",
            "'Scores'!$A$1",
            metadata={
                "hidden": True,
                "comment": "name comment",
                "description": "name description",
                "help": "name help",
                "statusBar": "name status",
            },
        )
        M.excel_set_document_properties(
            session_key,
            core={
                "contentStatus": "Draft",
                "identifier": "doc-id",
                "language": "vi-VN",
                "lastPrinted": "2025-01-03T04:05:06Z",
                "modified": "2025-01-02T03:04:05Z",
                "revision": "42",
                "version": "2.5",
            },
            modified_policy="set_explicit",
        )
        M.excel_set_header_footer(
            session_key,
            "Scores",
            {
                "odd_header": {"center": "ODD HEADER"},
                "odd_footer": {"center": "ODD FOOTER"},
                "even_header": {"left": "EVEN HEADER"},
                "even_footer": {"right": "EVEN FOOTER"},
                "first_header": {"center": "FIRST HEADER"},
                "first_footer": {"left": "FIRST FOOTER"},
            },
            {
                "alignWithMargins": False,
                "differentFirst": True,
                "differentOddEven": True,
                "scaleWithDoc": False,
            },
        )
        M.excel_set_hyperlink(
            session_key,
            "Scores",
            "A3",
            target="https://example.com/audit",
            display="Friendly display",
            tooltip="Audit tip",
        )
        M.excel_set_ignored_errors(
            session_key,
            "Scores",
            [{"sqref": "A1:A3", "numberStoredAsText": True}],
            mode="replace",
        )
        M.excel_set_sheet_properties(
            session_key,
            "Scores",
            {
                "codeName": "AuditSheetCode",
                "filterMode": True,
                "published": True,
                "syncRef": "D5",
                "outline": {
                    "applyStyles": True,
                    "summaryBelow": False,
                    "summaryRight": False,
                },
                "page_setup_properties": {
                    "autoPageBreaks": True,
                    "fitToPage": True,
                },
            },
        )
        M.excel_set_page_setup(
            session_key,
            "Scores",
            {
                "blackAndWhite": True,
                "cellComments": "asDisplayed",
                "copies": 2,
                "draft": True,
                "errors": "dash",
                "firstPageNumber": 7,
                "horizontalDpi": 300,
                "pageOrder": "overThenDown",
                "useFirstPageNumber": True,
                "verticalDpi": 300,
            },
            present=True,
            exact=True,
        )
        M.excel_set_print_options(
            session_key,
            "Scores",
            {
                "gridLines": True,
                "headings": True,
                "horizontalCentered": True,
                "verticalCentered": True,
            },
            present=True,
            exact=True,
        )
        M.excel_set_protected_ranges(
            session_key,
            "Scores",
            [{"name": "EditableRange", "sqref": "A1:A2", "password": "ABCD"}],
            mode="replace",
        )
        M.excel_set_row_properties(
            session_key,
            "Scores",
            1,
            {"collapsed": True, "thickBot": True, "thickTop": True},
        )
        M.excel_set_page_breaks(
            session_key,
            "Scores",
            row_breaks=[{"id": 10, "min": 0, "max": 16383, "man": True}],
            column_breaks=[{"id": 3, "min": 0, "max": 1048575, "man": True}],
        )
        M.excel_set_sheet_views(
            session_key,
            "Scores",
            [{"workbookViewId": 0, "selections": [{"activeCell": "D5", "sqref": "D5:E6"}]}],
        )
        M.excel_set_sheet_state(session_key, "Hidden", "hidden")
        M.excel_set_sheet_state(session_key, "VeryHidden", "veryHidden")
        M.excel_set_style(
            session_key,
            "Scores",
            0,
            0,
            style={
                "alignment": {
                    "justifyLastLine": True,
                    "readingOrder": 2,
                    "relativeIndent": 2,
                }
            },
        )
        M.excel_set_borders(
            session_key,
            "Scores",
            0,
            0,
            0,
            0,
            border={
                "outline": False,
                "start": {"style": "thin"},
                "end": {"style": "double"},
                "horizontal": {"style": "dotted"},
                "vertical": {"style": "dashed"},
            },
        )
        M.excel_set_cell_style_semantics(
            session_key,
            "Scores",
            "A1",
            xf={"applyBorder": True, "pivotButton": None, "quotePrefix": None},
        )
        M.excel_set_cell_style_semantics(
            session_key,
            "Scores",
            "A2",
            xf={"applyFont": True, "xfId": 1, "pivotButton": None, "quotePrefix": None},
            named_style="AuditStyle",
        )
        M.excel_set_workbook_properties(
            session_key,
            {
                "codeName": "AuditWorkbook",
                "date1904": True,
                "filterPrivacy": True,
                "saveExternalLinkValues": False,
                "showObjects": "none",
                "updateLinks": "never",
            },
            date_system_policy="preserve_displayed_dates",
        )
        M.excel_set_workbook_protection(
            session_key,
            {"lockStructure": True, "workbookPassword": "ABCD"},
            already_hashed=True,
        )
        M.excel_set_workbook_views(
            session_key,
            [
                {"activeTab": 0},
                {
                    "activeTab": 0,
                    "windowHeight": 8000,
                    "windowWidth": 12000,
                    "xWindow": 100,
                    "yWindow": 100,
                },
            ],
        )
        M.excel_save(session_key)
    finally:
        M.excel_close(session_key)


def _audit_87_status(path: Path) -> dict[str, bool]:
    with zipfile.ZipFile(path, "r") as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        worksheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        styles = ET.fromstring(archive.read("xl/styles.xml"))
        core = ET.fromstring(archive.read("docProps/core.xml"))
        sheet_relationships = ET.fromstring(
            archive.read("xl/worksheets/_rels/sheet1.xml.rels")
        )

    workbook_properties = _child(workbook, "workbookPr")
    workbook_protection = _child(workbook, "workbookProtection")
    calculation = _child(workbook, "calcPr")
    workbook_views = _child(workbook, "bookViews")
    sheets = _child(workbook, "sheets")
    defined_names = _child(workbook, "definedNames")
    audit_name = _node_by_attr(defined_names, "definedName", "name", "AuditName")

    sheet_properties = _child(worksheet, "sheetPr")
    outline = _child(sheet_properties, "outlinePr")
    page_setup_properties = _child(sheet_properties, "pageSetUpPr")
    sheet_views = _child(worksheet, "sheetViews")
    first_sheet_view = _child(sheet_views, "sheetView")
    selection = _child(first_sheet_view, "selection")
    page_setup = _child(worksheet, "pageSetup")
    print_options = _child(worksheet, "printOptions")
    header_footer = _child(worksheet, "headerFooter")
    row_breaks = _child(worksheet, "rowBreaks")
    column_breaks = _child(worksheet, "colBreaks")
    row_break = _child(row_breaks, "brk")
    column_break = _child(column_breaks, "brk")
    protected_range = _child(_child(worksheet, "protectedRanges"), "protectedRange")
    ignored_error = _child(_child(worksheet, "ignoredErrors"), "ignoredError")
    auto_filter = _child(worksheet, "autoFilter")
    filter_column = _child(auto_filter, "filterColumn")
    filter_value = _child(_child(filter_column, "filters"), "filter")
    sort_state = _child(auto_filter, "sortState")
    sort_condition = _child(sort_state, "sortCondition")
    hyperlink = _child(_child(worksheet, "hyperlinks"), "hyperlink")
    hyperlink_relationship = next(
        (
            item
            for item in sheet_relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
            if hyperlink is not None and item.get("Id") == hyperlink.get(REL_ID)
        ),
        None,
    )

    rows = _child(worksheet, "sheetData")
    row2 = _node_by_attr(rows, "row", "r", "2")
    cell_a1 = worksheet.find(f".//{MAIN}c[@r='A1']")
    cell_a2 = worksheet.find(f".//{MAIN}c[@r='A2']")
    cell_b2 = worksheet.find(f".//{MAIN}c[@r='B2']")
    formula_b2_value = _child(cell_b2, "v")

    cell_xfs_node = _child(styles, "cellXfs")
    cell_xfs = list(cell_xfs_node) if cell_xfs_node is not None else []
    a1_xf = cell_xfs[int(cell_a1.get("s", "0"))]
    a2_xf = cell_xfs[int(cell_a2.get("s", "0"))]
    a1_alignment = _child(a1_xf, "alignment")
    borders_node = _child(styles, "borders")
    borders = list(borders_node) if borders_node is not None else []
    a1_border_id = int(a1_xf.get("borderId", "0"))
    a1_border = borders[a1_border_id]
    style_names = {
        item.get("name") for item in _children(_child(styles, "cellStyles"), "cellStyle")
    }
    sheet_states = {
        item.get("name"): item.get("state", "visible")
        for item in _children(sheets, "sheet")
    }
    core_values = {_local_name(item.tag): item.text for item in list(core)}
    header_texts = {
        name: (_child(header_footer, name).text if _child(header_footer, name) is not None else None)
        for name in (
            "oddHeader",
            "oddFooter",
            "evenHeader",
            "evenFooter",
            "firstHeader",
            "firstFooter",
        )
    }

    status = {
        "autoFilter.children": (
            _attrs(filter_column).get("colId") == "0"
            and _attrs(filter_value).get("val") == "Alice"
            and _attrs(sort_state).get("ref") == "A2:B3"
            and _attrs(sort_condition).get("ref") == "B2:B3"
            and _attrs(sort_condition).get("descending") == "1"
        ),
        "calcPr.calcMode": _attrs(calculation).get("calcMode") == "manual",
        "calcPr.forceFullCalc": _attrs(calculation).get("forceFullCalc") == "0",
        "calcPr.fullCalcOnLoad": _attrs(calculation).get("fullCalcOnLoad") == "0",
        "calcPr.iterate": _attrs(calculation).get("iterate") == "1",
        "calcPr.iterateCount": _attrs(calculation).get("iterateCount") == "55",
        "calcPr.iterateDelta": _attrs(calculation).get("iterateDelta") == "0.002",
        "cell_style_names": {"Normal", "AuditStyle"}.issubset(style_names),
        "colBreaks.attrs.count": _attrs(column_breaks).get("count") == "1",
        "colBreaks.attrs.manualBreakCount": _attrs(column_breaks).get("manualBreakCount") == "1",
        "colBreaks.children": _attrs(column_break) == {
            "id": "3",
            "min": "0",
            "max": "1048575",
            "man": "1",
        },
        "core_props.contentStatus": core_values.get("contentStatus") == "Draft",
        "core_props.identifier": core_values.get("identifier") == "doc-id",
        "core_props.language": core_values.get("language") == "vi-VN",
        "core_props.lastPrinted": core_values.get("lastPrinted") == "2025-01-03T04:05:06Z",
        "core_props.modified": core_values.get("modified") == "2025-01-02T03:04:05Z",
        "core_props.revision": core_values.get("revision") == "42",
        "core_props.version": core_values.get("version") == "2.5",
        "defined_names": (
            audit_name is not None
            and (audit_name.text or "") == "'Scores'!$A$1"
            and _attrs(audit_name) == {
                "name": "AuditName",
                "comment": "name comment",
                "description": "name description",
                "help": "name help",
                "statusBar": "name status",
                "hidden": "1",
            }
        ),
        "formula_B2.v": formula_b2_value is not None and formula_b2_value.text == "10",
        "headerFooter.attrs.alignWithMargins": _attrs(header_footer).get("alignWithMargins") == "0",
        "headerFooter.attrs.differentFirst": _attrs(header_footer).get("differentFirst") == "1",
        "headerFooter.attrs.differentOddEven": _attrs(header_footer).get("differentOddEven") == "1",
        "headerFooter.attrs.scaleWithDoc": _attrs(header_footer).get("scaleWithDoc") == "0",
        "headerFooter.children": header_texts == {
            "oddHeader": "&CODD HEADER",
            "oddFooter": "&CODD FOOTER",
            "evenHeader": "&LEVEN HEADER",
            "evenFooter": "&REVEN FOOTER",
            "firstHeader": "&CFIRST HEADER",
            "firstFooter": "&LFIRST FOOTER",
        },
        "hyperlinks.children": (
            _attrs(hyperlink).get("ref") == "A3"
            and _attrs(hyperlink).get("display") == "Friendly display"
            and _attrs(hyperlink).get("tooltip") == "Audit tip"
            and hyperlink_relationship is not None
            and hyperlink_relationship.get("Target") == "https://example.com/audit"
            and hyperlink_relationship.get("TargetMode") == "External"
        ),
        "ignoredErrors.children": _attrs(ignored_error) == {
            "sqref": "A1:A3",
            "numberStoredAsText": "1",
        },
        "outlinePr.applyStyles": _attrs(outline).get("applyStyles") == "1",
        "outlinePr.summaryBelow": _attrs(outline).get("summaryBelow") == "0",
        "outlinePr.summaryRight": _attrs(outline).get("summaryRight") == "0",
        "pageSetUpPr.autoPageBreaks": _attrs(page_setup_properties).get("autoPageBreaks") == "1",
        "pageSetup.attrs.blackAndWhite": _attrs(page_setup).get("blackAndWhite") == "1",
        "pageSetup.attrs.cellComments": _attrs(page_setup).get("cellComments") == "asDisplayed",
        "pageSetup.attrs.copies": _attrs(page_setup).get("copies") == "2",
        "pageSetup.attrs.draft": _attrs(page_setup).get("draft") == "1",
        "pageSetup.attrs.errors": _attrs(page_setup).get("errors") == "dash",
        "pageSetup.attrs.firstPageNumber": _attrs(page_setup).get("firstPageNumber") == "7",
        "pageSetup.attrs.horizontalDpi": _attrs(page_setup).get("horizontalDpi") == "300",
        "pageSetup.attrs.pageOrder": _attrs(page_setup).get("pageOrder") == "overThenDown",
        "pageSetup.attrs.useFirstPageNumber": _attrs(page_setup).get("useFirstPageNumber") == "1",
        "pageSetup.attrs.verticalDpi": _attrs(page_setup).get("verticalDpi") == "300",
        "pageSetup.children": page_setup is not None,
        "printOptions.attrs.gridLines": _attrs(print_options).get("gridLines") == "1",
        "printOptions.attrs.headings": _attrs(print_options).get("headings") == "1",
        "printOptions.attrs.horizontalCentered": _attrs(print_options).get("horizontalCentered") == "1",
        "printOptions.attrs.verticalCentered": _attrs(print_options).get("verticalCentered") == "1",
        "printOptions.children": print_options is not None,
        "protectedRanges.children": _attrs(protected_range) == {
            "name": "EditableRange",
            "sqref": "A1:A2",
            "password": "ABCD",
        },
        "row2.collapsed": _attrs(row2).get("collapsed") == "1",
        "row2.thickBot": _attrs(row2).get("thickBot") == "1",
        "row2.thickTop": _attrs(row2).get("thickTop") == "1",
        "rowBreaks.attrs.count": _attrs(row_breaks).get("count") == "1",
        "rowBreaks.attrs.manualBreakCount": _attrs(row_breaks).get("manualBreakCount") == "1",
        "rowBreaks.children": _attrs(row_break) == {
            "id": "10",
            "min": "0",
            "max": "16383",
            "man": "1",
        },
        "selection.activeCell": _attrs(selection).get("activeCell") == "D5",
        "selection.sqref": _attrs(selection).get("sqref") == "D5:E6",
        "sheetPr.codeName": _attrs(sheet_properties).get("codeName") == "AuditSheetCode",
        "sheetPr.filterMode": _attrs(sheet_properties).get("filterMode") == "1",
        "sheetPr.published": _attrs(sheet_properties).get("published") == "1",
        "sheetPr.syncRef": _attrs(sheet_properties).get("syncRef") == "D5",
        "sheet_states.Hidden": sheet_states.get("Hidden") == "hidden",
        "sheet_states.VeryHidden": sheet_states.get("VeryHidden") == "veryHidden",
        "style_A1.alignment.justifyLastLine": _attrs(a1_alignment).get("justifyLastLine") == "1",
        "style_A1.alignment.readingOrder": _attrs(a1_alignment).get("readingOrder") == "2",
        "style_A1.alignment.relativeIndent": _attrs(a1_alignment).get("relativeIndent") == "2",
        "style_A1.border.outline": _attrs(a1_border).get("outline") == "0",
        "style_A1.border_sides.end.style": _attrs(_child(a1_border, "end")).get("style") == "double",
        "style_A1.border_sides.horizontal.style": _attrs(_child(a1_border, "horizontal")).get("style") == "dotted",
        "style_A1.border_sides.start.style": _attrs(_child(a1_border, "start")).get("style") == "thin",
        "style_A1.border_sides.vertical.style": _attrs(_child(a1_border, "vertical")).get("style") == "dashed",
        "style_A1.xf.applyBorder": _attrs(a1_xf).get("applyBorder") == "1",
        "style_A1.xf.borderId": a1_border_id > 0,
        "style_A1.xf.pivotButton": "pivotButton" not in _attrs(a1_xf),
        "style_A1.xf.quotePrefix": "quotePrefix" not in _attrs(a1_xf),
        "style_A2.xf.applyFont": _attrs(a2_xf).get("applyFont") == "1",
        "style_A2.xf.pivotButton": "pivotButton" not in _attrs(a2_xf),
        "style_A2.xf.quotePrefix": "quotePrefix" not in _attrs(a2_xf),
        "style_A2.xf.xfId": _attrs(a2_xf).get("xfId") == "1",
        "workbookPr.codeName": _attrs(workbook_properties).get("codeName") == "AuditWorkbook",
        "workbookPr.date1904": _attrs(workbook_properties).get("date1904") == "1",
        "workbookPr.filterPrivacy": _attrs(workbook_properties).get("filterPrivacy") == "1",
        "workbookPr.saveExternalLinkValues": _attrs(workbook_properties).get("saveExternalLinkValues") == "0",
        "workbookPr.showObjects": _attrs(workbook_properties).get("showObjects") == "none",
        "workbookPr.updateLinks": _attrs(workbook_properties).get("updateLinks") == "never",
        "workbookProtection.lockStructure": _attrs(workbook_protection).get("lockStructure") == "1",
        "workbookProtection.workbookPassword": _attrs(workbook_protection).get("workbookPassword") == "ABCD",
        "workbook_view_count": len(_children(workbook_views, "workbookView")) == 2,
    }
    assert len(status) == 87
    return status


def _exact_87_report(path: Path) -> dict:
    status = _audit_87_status(path)
    failed = [key for key, preserved in status.items() if not preserved]
    return {
        "path": str(path),
        "checked_key_count": len(status),
        "preserved_key_count": len(status) - len(failed),
        "failed_key_count": len(failed),
        "failed_keys": failed,
    }


def _assert_exact_87(path: Path) -> dict:
    report = _exact_87_report(path)
    assert report["checked_key_count"] == 87
    assert report["preserved_key_count"] == 87
    assert report["failed_key_count"] == 0
    assert report["failed_keys"] == [], f"Exact-87 failures for {path}: {report['failed_keys']}"
    return report


def test_public_tools_create_reload_and_preserve_exact_87(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCLOUPE_EXCEL_BACKUP_DIR", str(tmp_path / "backups"))
    output_path = tmp_path / "public-exact-87.xlsx"
    roundtrip_path = tmp_path / "public-exact-87-roundtrip.xlsx"

    _create_exact87_workbook(output_path)
    assert json.loads(M.excel_validate_workbook(str(output_path)))["valid"] is True
    created_report = _assert_exact_87(output_path)
    assert created_report == {
        "path": str(output_path),
        "checked_key_count": 87,
        "preserved_key_count": 87,
        "failed_key_count": 0,
        "failed_keys": [],
    }

    session_key = _load_key(M.excel_load(str(output_path)))
    try:
        M.excel_save_as_copy(session_key, str(roundtrip_path))
    finally:
        M.excel_close(session_key)

    assert json.loads(M.excel_validate_workbook(str(roundtrip_path)))["valid"] is True
    _assert_exact_87(roundtrip_path)
    verification = json.loads(
        M.excel_verify_preservation(
            after_path=str(roundtrip_path),
            before_path=str(output_path),
            max_differences=5000,
        )
    )
    assert verification["preservation_ok"] is True
    assert verification["unapproved_difference_count"] == 0


def test_exact_87_checker_matches_legacy_reference_when_available():
    if not LEGACY_SOURCE.is_file():
        pytest.skip(f"Legacy exact-87 source is unavailable: {LEGACY_SOURCE}")
    _assert_exact_87(LEGACY_SOURCE)
