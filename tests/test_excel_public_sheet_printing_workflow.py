"""Public save/reload acceptance coverage for worksheet and printing semantics."""

import json
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers" / "excel"))

import main as M  # noqa: E402


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"main": SHEET_NS}


def _load_key(load_result: str) -> str:
    return load_result.split("session_key='")[1].split("'")[0]


def _new_workbook(tmp_path: Path, track_session, sheet_names=None):
    output_path = tmp_path / "sheet-printing.xlsx"
    created = json.loads(
        M.excel_create_workbook(
            sheet_names=sheet_names or ["Report"],
            target_path=str(output_path),
        )
    )
    return track_session(created["session_key"]), output_path


def _roundtrip_semantics(track_session, session_key: str, path: Path, *sheet_names: str):
    M.excel_save(session_key)
    M.excel_close(session_key)
    reloaded_key = track_session(_load_key(M.excel_load(str(path))))
    try:
        return {
            sheet_name: json.loads(M.excel_get_sheet_semantics(reloaded_key, sheet_name))
            for sheet_name in sheet_names
        }
    finally:
        M.excel_close(reloaded_key)


def _worksheet_root(path: Path, index: int = 1):
    with zipfile.ZipFile(path, "r") as archive:
        return ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))


def _workbook_root(path: Path):
    with zipfile.ZipFile(path, "r") as archive:
        return ET.fromstring(archive.read("xl/workbook.xml"))


@pytest.fixture
def track_session():
    session_keys = []

    def track(session_key: str) -> str:
        session_keys.append(session_key)
        return session_key

    yield track

    for session_key in reversed(session_keys):
        try:
            M.excel_close(session_key)
        except Exception:
            pass


def test_set_sheet_state_roundtrips_and_guards_last_visible_sheet(tmp_path, track_session):
    session_key, path = _new_workbook(
        tmp_path,
        track_session,
        sheet_names=["Visible", "Hidden", "VeryHidden"],
    )

    M.excel_set_sheet_state(session_key, "Hidden", "hidden")
    M.excel_set_sheet_state(session_key, "VeryHidden", "veryHidden")
    with pytest.raises(ValueError, match="At least one worksheet must remain visible"):
        M.excel_set_sheet_state(session_key, "Visible", "hidden")

    semantics = _roundtrip_semantics(
        track_session,
        session_key,
        path,
        "Visible",
        "Hidden",
        "VeryHidden",
    )
    assert {name: value["state"] for name, value in semantics.items()} == {
        "Visible": "visible",
        "Hidden": "hidden",
        "VeryHidden": "veryHidden",
    }

    workbook = openpyxl.load_workbook(path)
    try:
        assert {sheet.title: sheet.sheet_state for sheet in workbook.worksheets} == {
            "Visible": "visible",
            "Hidden": "hidden",
            "VeryHidden": "veryHidden",
        }
    finally:
        workbook.close()


def test_set_sheet_properties_roundtrips_structured_children(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    expected = {
        "codeName": "ReportCode",
        "filterMode": True,
        "published": True,
        "syncRef": "C3",
        "syncHorizontal": True,
        "syncVertical": False,
        "outline": {
            "applyStyles": True,
            "summaryBelow": False,
            "summaryRight": True,
            "showOutlineSymbols": False,
        },
        "page_setup_properties": {
            "autoPageBreaks": False,
            "fitToPage": True,
        },
    }
    M.excel_set_sheet_properties(
        session_key,
        "Report",
        {**expected, "published": False},
    )
    M.excel_set_sheet_properties(session_key, "Report", {"published": True})

    semantics = _roundtrip_semantics(track_session, session_key, path, "Report")["Report"]
    assert semantics["sheet_properties"] == expected

    root = _worksheet_root(path)
    sheet_pr = root.find("main:sheetPr", NS)
    assert sheet_pr is not None
    assert {key: sheet_pr.attrib.get(key) for key in (
        "codeName",
        "filterMode",
        "published",
        "syncRef",
        "syncHorizontal",
        "syncVertical",
    )} == {
        "codeName": "ReportCode",
        "filterMode": "1",
        "published": "1",
        "syncRef": "C3",
        "syncHorizontal": "1",
        "syncVertical": "0",
    }
    outline = sheet_pr.find("main:outlinePr", NS)
    assert outline is not None
    assert outline.attrib == {
        "applyStyles": "1",
        "summaryBelow": "0",
        "summaryRight": "1",
        "showOutlineSymbols": "0",
    }
    page_setup_properties = sheet_pr.find("main:pageSetUpPr", NS)
    assert page_setup_properties is not None
    assert page_setup_properties.attrib == {
        "autoPageBreaks": "0",
        "fitToPage": "1",
    }


def test_sheet_views_roundtrip_multiple_views_pane_and_selections_without_a1_reset(
    tmp_path,
    track_session,
):
    session_key, path = _new_workbook(tmp_path, track_session)
    views = [
        {
            "workbookViewId": 0,
            "showGridLines": False,
            "zoomScale": 90,
            "topLeftCell": "C3",
            "pane": {
                "xSplit": 2,
                "ySplit": 3,
                "topLeftCell": "C4",
                "activePane": "bottomRight",
                "state": "frozenSplit",
            },
            "selections": [
                {"pane": "topRight", "activeCell": "D4", "sqref": "D4 E5"},
                {"pane": "bottomRight", "activeCell": "F6", "sqref": "F6:G7"},
            ],
        },
        {
            "workbookViewId": 0,
            "tabSelected": False,
            "topLeftCell": "H8",
            "selections": [
                {"activeCell": "H8", "sqref": "H8:I9"},
            ],
        },
    ]
    M.excel_set_sheet_views(session_key, "Report", views)
    assert json.loads(M.excel_get_sheet_views(session_key, "Report"))["views"] == views

    _roundtrip_semantics(track_session, session_key, path, "Report")
    reloaded_key = track_session(_load_key(M.excel_load(str(path))))
    try:
        reloaded_views = json.loads(M.excel_get_sheet_views(reloaded_key, "Report"))["views"]
        for view in reloaded_views:
            view.pop("zoom", None)
        assert reloaded_views == views
    finally:
        M.excel_close(reloaded_key)

    root = _worksheet_root(path)
    disk_views = root.findall("main:sheetViews/main:sheetView", NS)
    assert len(disk_views) == 2
    pane = disk_views[0].find("main:pane", NS)
    assert pane is not None
    assert pane.attrib == {
        "xSplit": "2",
        "ySplit": "3",
        "topLeftCell": "C4",
        "activePane": "bottomRight",
        "state": "frozenSplit",
    }
    selections = [selection.attrib for selection in disk_views[0].findall("main:selection", NS)]
    assert selections == [
        {"pane": "topRight", "activeCell": "D4", "sqref": "D4 E5"},
        {"pane": "bottomRight", "activeCell": "F6", "sqref": "F6:G7"},
    ]
    assert all(selection.get("activeCell") != "A1" for view in disk_views for selection in view.findall("main:selection", NS))


def test_set_row_properties_partial_patch_preserves_other_flags(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    M.excel_set_row_properties(
        session_key,
        "Report",
        0,
        {
            "height": 31.5,
            "hidden": True,
            "outlineLevel": 3,
            "collapsed": True,
            "thickTop": True,
            "thickBot": True,
            "customFormat": True,
            "customHeight": True,
            "style": 0,
            "phonetic": True,
        },
    )
    M.excel_set_row_properties(session_key, "Report", 0, {"height": 32.25})

    _roundtrip_semantics(track_session, session_key, path, "Report")
    row = _worksheet_root(path).find("main:sheetData/main:row", NS)
    assert row is not None
    assert {key: row.attrib.get(key) for key in (
        "ht",
        "hidden",
        "outlineLevel",
        "collapsed",
        "thickTop",
        "thickBot",
        "customFormat",
        "customHeight",
        "s",
        "ph",
    )} == {
        "ht": "32.25",
        "hidden": "1",
        "outlineLevel": "3",
        "collapsed": "1",
        "thickTop": "1",
        "thickBot": "1",
        "customFormat": "1",
        "customHeight": "1",
        "s": "0",
        "ph": "1",
    }


def test_set_page_setup_roundtrips_full_attribute_set(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    properties = {
        "blackAndWhite": True,
        "cellComments": "asDisplayed",
        "copies": 2,
        "draft": True,
        "errors": "dash",
        "firstPageNumber": 3,
        "horizontalDpi": 300,
        "verticalDpi": 600,
        "pageOrder": "overThenDown",
        "useFirstPageNumber": True,
        "orientation": "landscape",
        "paperSize": "9",
        "scale": 85,
        "fitToWidth": 2,
        "fitToHeight": 4,
        "fitToPage": True,
    }
    M.excel_set_page_setup(session_key, "Report", properties, present=True, exact=True)

    _roundtrip_semantics(track_session, session_key, path, "Report")
    root = _worksheet_root(path)
    page_setup = root.find("main:pageSetup", NS)
    assert page_setup is not None
    assert {key: page_setup.attrib.get(key) for key in properties if key != "fitToPage"} == {
        "blackAndWhite": "1",
        "cellComments": "asDisplayed",
        "copies": "2",
        "draft": "1",
        "errors": "dash",
        "firstPageNumber": "3",
        "horizontalDpi": "300",
        "verticalDpi": "600",
        "pageOrder": "overThenDown",
        "useFirstPageNumber": "1",
        "orientation": "landscape",
        "paperSize": "9",
        "scale": "85",
        "fitToWidth": "2",
        "fitToHeight": "4",
    }
    page_setup_properties = root.find("main:sheetPr/main:pageSetUpPr", NS)
    assert page_setup_properties is not None
    assert page_setup_properties.attrib.get("fitToPage") == "1"


def test_set_page_setup_exact_mode_preserves_explicit_empty_element(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    M.excel_set_page_setup(
        session_key,
        "Report",
        {"orientation": "portrait", "copies": 2},
        present=True,
    )
    M.excel_set_page_setup(session_key, "Report", {}, present=True, exact=True)

    _roundtrip_semantics(track_session, session_key, path, "Report")
    page_setup = _worksheet_root(path).find("main:pageSetup", NS)
    assert page_setup is not None
    assert page_setup.attrib == {}


def test_set_print_options_roundtrips_all_explicit_attributes(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    properties = {
        "gridLines": False,
        "gridLinesSet": True,
        "headings": True,
        "horizontalCentered": False,
        "verticalCentered": True,
    }
    M.excel_set_print_options(session_key, "Report", properties, present=True, exact=True)

    _roundtrip_semantics(track_session, session_key, path, "Report")
    print_options = _worksheet_root(path).find("main:printOptions", NS)
    assert print_options is not None
    assert print_options.attrib == {
        "gridLines": "0",
        "gridLinesSet": "1",
        "headings": "1",
        "horizontalCentered": "0",
        "verticalCentered": "1",
    }


def test_set_print_options_exact_mode_preserves_explicit_empty_element(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    M.excel_set_print_options(session_key, "Report", {"gridLines": True}, present=True)
    M.excel_set_print_options(session_key, "Report", {}, present=True, exact=True)

    _roundtrip_semantics(track_session, session_key, path, "Report")
    print_options = _worksheet_root(path).find("main:printOptions", NS)
    assert print_options is not None
    assert print_options.attrib == {}


def test_set_header_footer_partial_patch_preserves_every_other_section(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    sections = {
        "odd_header": {"left": "Odd Header Left", "center": "Odd Header Center", "right": "Odd Header Right"},
        "odd_footer": {"left": "Odd Footer Left", "center": "Odd Footer Center", "right": "Odd Footer Right"},
        "even_header": {"left": "Even Header Left", "center": "Even Header Center", "right": "Even Header Right"},
        "even_footer": {"left": "Even Footer Left", "center": "Even Footer Center", "right": "Even Footer Right"},
        "first_header": {"left": "First Header Left", "center": "First Header Center", "right": "First Header Right"},
        "first_footer": {"left": "First Footer Left", "center": "First Footer Center", "right": "First Footer Right"},
    }
    properties = {
        "alignWithMargins": False,
        "differentFirst": True,
        "differentOddEven": True,
        "scaleWithDoc": False,
    }
    M.excel_set_header_footer(session_key, "Report", sections, properties)
    M.excel_set_header_footer(
        session_key,
        "Report",
        {"odd_header": {"center": "Odd Header Center Patched"}},
    )

    _roundtrip_semantics(track_session, session_key, path, "Report")
    header_footer = _worksheet_root(path).find("main:headerFooter", NS)
    assert header_footer is not None
    assert header_footer.attrib == {
        "alignWithMargins": "0",
        "differentFirst": "1",
        "differentOddEven": "1",
        "scaleWithDoc": "0",
    }
    expected_text = {
        "oddHeader": ("Odd Header Left", "Odd Header Center Patched", "Odd Header Right"),
        "oddFooter": ("Odd Footer Left", "Odd Footer Center", "Odd Footer Right"),
        "evenHeader": ("Even Header Left", "Even Header Center", "Even Header Right"),
        "evenFooter": ("Even Footer Left", "Even Footer Center", "Even Footer Right"),
        "firstHeader": ("First Header Left", "First Header Center", "First Header Right"),
        "firstFooter": ("First Footer Left", "First Footer Center", "First Footer Right"),
    }
    for tag, pieces in expected_text.items():
        text = header_footer.find(f"main:{tag}", NS)
        assert text is not None
        assert all(piece in (text.text or "") for piece in pieces)
    assert "Odd Header Center&" not in (header_footer.find("main:oddHeader", NS).text or "")


def test_set_page_breaks_roundtrips_counters_and_break_attributes(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    M.excel_set_page_breaks(
        session_key,
        "Report",
        row_breaks=[
            {"id": 4, "min": 0, "max": 16383, "man": True, "pt": False},
            {"id": 9, "min": 1, "max": 20, "man": False, "pt": True},
        ],
        column_breaks=[
            {"id": 2, "min": 0, "max": 1048575, "man": True, "pt": True},
        ],
    )

    _roundtrip_semantics(track_session, session_key, path, "Report")
    root = _worksheet_root(path)
    row_breaks = root.find("main:rowBreaks", NS)
    column_breaks = root.find("main:colBreaks", NS)
    assert row_breaks is not None
    assert column_breaks is not None
    assert row_breaks.attrib == {"count": "2", "manualBreakCount": "1"}
    assert column_breaks.attrib == {"count": "1", "manualBreakCount": "1"}
    assert [item.attrib for item in row_breaks.findall("main:brk", NS)] == [
        {"id": "4", "min": "0", "max": "16383", "man": "1", "pt": "0"},
        {"id": "9", "min": "1", "max": "20", "man": "0", "pt": "1"},
    ]
    assert [item.attrib for item in column_breaks.findall("main:brk", NS)] == [
        {"id": "2", "min": "0", "max": "1048575", "man": "1", "pt": "1"},
    ]


def test_set_print_area_and_titles_roundtrip_multi_area_rows_and_columns(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    M.excel_set_print_area(
        session_key,
        "Report",
        ["$A$1:$B$4", "$D$2:$F$9"],
    )
    M.excel_set_print_titles(
        session_key,
        "Report",
        repeated_rows="$1:$3",
        repeated_columns="$A:$C",
    )

    semantics = _roundtrip_semantics(track_session, session_key, path, "Report")["Report"]
    assert "$A$1:$B$4" in semantics["print_area"]
    assert "$D$2:$F$9" in semantics["print_area"]
    assert semantics["print_titles"] == {"rows": "$1:$3", "cols": "$A:$C"}

    workbook = openpyxl.load_workbook(path)
    try:
        worksheet = workbook["Report"]
        print_area = str(worksheet.print_area).replace(" ", "")
        assert "$A$1:$B$4" in print_area
        assert "$D$2:$F$9" in print_area
        assert worksheet.print_title_rows == "$1:$3"
        assert worksheet.print_title_cols == "$A:$C"
    finally:
        workbook.close()

    defined_names = _workbook_root(path).findall("main:definedNames/main:definedName", NS)
    by_name = {item.attrib["name"]: item for item in defined_names}
    assert by_name["_xlnm.Print_Area"].attrib.get("localSheetId") == "0"
    assert "$A$1:$B$4" in (by_name["_xlnm.Print_Area"].text or "")
    assert "$D$2:$F$9" in (by_name["_xlnm.Print_Area"].text or "")
    assert by_name["_xlnm.Print_Titles"].attrib.get("localSheetId") == "0"
    assert "$A:$C" in (by_name["_xlnm.Print_Titles"].text or "")
    assert "$1:$3" in (by_name["_xlnm.Print_Titles"].text or "")


def test_set_protected_ranges_roundtrips_without_enabling_sheet_protection(tmp_path, track_session):
    session_key, path = _new_workbook(tmp_path, track_session)
    M.excel_set_protected_ranges(
        session_key,
        "Report",
        [
            {"name": "InputCells", "sqref": "B2:C4", "password": "ABCD"},
            {"name": "ReviewCells", "sqref": "E2 E4", "securityDescriptor": "S-1-5-21"},
        ],
    )
    M.excel_set_protected_ranges(
        session_key,
        "Report",
        [
            {"name": "InputCells", "sqref": "B2:D4", "password": "ABCD"},
            {"name": "ReviewCells", "delete": True},
            {
                "name": "ApprovalCells",
                "sqref": "F6:G7",
                "algorithmName": "SHA-512",
                "hashValue": "AA==",
                "saltValue": "AQ==",
                "spinCount": 100000,
            },
        ],
        mode="patch",
    )

    semantics = _roundtrip_semantics(track_session, session_key, path, "Report")["Report"]
    assert semantics["protection"] is None
    assert semantics["protected_ranges"] == [
        {"name": "InputCells", "sqref": "B2:D4", "password": "ABCD"},
        {
            "name": "ApprovalCells",
            "sqref": "F6:G7",
            "algorithmName": "SHA-512",
            "hashValue": "AA==",
            "saltValue": "AQ==",
            "spinCount": 100000,
        },
    ]

    root = _worksheet_root(path)
    assert root.find("main:sheetProtection", NS) is None
    protected_ranges = root.findall("main:protectedRanges/main:protectedRange", NS)
    assert [item.attrib for item in protected_ranges] == [
        {"name": "InputCells", "sqref": "B2:D4", "password": "ABCD"},
        {
            "name": "ApprovalCells",
            "sqref": "F6:G7",
            "algorithmName": "SHA-512",
            "hashValue": "AA==",
            "saltValue": "AQ==",
            "spinCount": "100000",
        },
    ]


def test_set_sheet_protection_partial_patch_preserves_siblings_and_explicit_false(
    tmp_path,
    track_session,
    monkeypatch,
):
    monkeypatch.setenv("DOCLOUPE_EXCEL_BACKUP_DIR", str(tmp_path / "backups"))
    session_key, path = _new_workbook(tmp_path, track_session)
    created = json.loads(
        M.excel_set_sheet_protection(
            session_key,
            "Report",
            {
                "password": "ABCD",
                "selectLockedCells": True,
                "selectUnlockedCells": False,
                "formatCells": False,
                "autoFilter": True,
                "objects": True,
            },
            enabled=True,
            already_hashed=True,
        )
    )
    assert created["after"]["password"] == "ABCD"
    assert created["after"]["password_is_hashed"] is True

    patched = json.loads(
        M.excel_set_sheet_protection(
            session_key,
            "Report",
            {"autoFilter": False, "formatRows": False},
        )
    )
    assert patched["after"]["password"] == "ABCD"
    assert patched["after"]["objects"] is True
    assert patched["after"]["autoFilter"] is False
    assert patched["after"]["formatRows"] is False

    semantics = _roundtrip_semantics(track_session, session_key, path, "Report")["Report"]
    protection = semantics["protection"]
    assert protection["password"] == "ABCD"
    assert protection["password_is_hashed"] is True
    assert protection["selectLockedCells"] is True
    assert protection["selectUnlockedCells"] is False
    assert protection["formatCells"] is False
    assert protection["autoFilter"] is False
    assert protection["objects"] is True
    assert protection["formatRows"] is False

    sheet_protection = _worksheet_root(path).find("main:sheetProtection", NS)
    assert sheet_protection is not None
    assert sheet_protection.attrib["sheet"] == "1"
    assert sheet_protection.attrib["password"] == "ABCD"
    assert sheet_protection.attrib["autoFilter"] == "0"
    assert sheet_protection.attrib["objects"] == "1"


def test_set_sheet_protection_modern_hash_null_clear_plaintext_and_disable(
    tmp_path,
    track_session,
    monkeypatch,
):
    monkeypatch.setenv("DOCLOUPE_EXCEL_BACKUP_DIR", str(tmp_path / "backups"))
    session_key, path = _new_workbook(tmp_path, track_session)
    M.excel_set_sheet_protection(
        session_key,
        "Report",
        {
            "algorithmName": "SHA-512",
            "hashValue": "AA==",
            "saltValue": "AQ==",
            "spinCount": 0,
            "sort": False,
        },
        enabled=True,
    )
    M.excel_save(session_key)
    M.excel_close(session_key)
    session_key = track_session(_load_key(M.excel_load(str(path))))
    protection = json.loads(M.excel_get_sheet_semantics(session_key, "Report"))["protection"]
    assert protection["algorithmName"] == "SHA-512"
    assert protection["hashValue"] == "AA=="
    assert protection["saltValue"] == "AQ=="
    assert protection["spinCount"] == 0
    assert protection["sort"] is False

    converted = json.loads(
        M.excel_set_sheet_protection(
            session_key,
            "Report",
            {
                "algorithmName": None,
                "hashValue": None,
                "saltValue": None,
                "spinCount": None,
                "password": "secret",
                "clear_nulls": True,
            },
            already_hashed=False,
        )
    )
    assert converted["after"]["password"] == "secret"
    assert converted["after"]["password_is_hashed"] is False
    assert "algorithmName" not in converted["after"]
    assert "hashValue" not in converted["after"]
    assert "saltValue" not in converted["after"]
    assert "spinCount" not in converted["after"]

    M.excel_save(session_key)
    M.excel_close(session_key)
    session_key = track_session(_load_key(M.excel_load(str(path))))
    reloaded = json.loads(M.excel_get_sheet_semantics(session_key, "Report"))["protection"]
    assert reloaded["password"] != "secret"
    assert reloaded["password_is_hashed"] is True
    assert "algorithmName" not in reloaded
    assert "hashValue" not in reloaded
    assert "saltValue" not in reloaded
    assert "spinCount" not in reloaded

    disabled = json.loads(
        M.excel_set_sheet_protection(session_key, "Report", enabled=False)
    )
    assert disabled["after"] is None
    M.excel_save(session_key)
    assert _worksheet_root(path).find("main:sheetProtection", NS) is None


def test_structural_delete_shifts_sheet_metadata_and_recomputes_break_counters(
    tmp_path,
    track_session,
):
    session_key, path = _new_workbook(tmp_path, track_session)
    M.excel_edit_cells(
        session_key,
        "Report",
        [{"row_index": 9, "edits": {0: 10}}],
    )
    M.excel_set_page_breaks(
        session_key,
        "Report",
        row_breaks=[
            {"id": 4, "min": 0, "max": 16383, "man": True},
            {"id": 9, "min": 0, "max": 16383, "man": False},
        ],
        column_breaks=[
            {"id": 2, "min": 0, "max": 1048575, "man": True},
        ],
    )
    M.excel_set_protected_ranges(
        session_key,
        "Report",
        [{"name": "Editable", "sqref": "B5:C7", "password": "ABCD"}],
    )
    M.excel_set_sheet_views(
        session_key,
        "Report",
        [
            {
                "workbookViewId": 0,
                "topLeftCell": "C5",
                "pane": {"topLeftCell": "C6", "state": "frozen", "ySplit": 5},
                "selections": [
                    {"activeCell": "D5", "sqref": "D5:E7"},
                ],
            },
        ],
    )
    M.excel_set_print_area(
        session_key,
        "Report",
        ["$A$2:$B$6", "$D$5:$E$9"],
    )
    M.excel_set_print_titles(
        session_key,
        "Report",
        repeated_rows="$1:$2",
    )

    M.excel_delete_rows(session_key, "Report", row_indices=[3])
    semantics = _roundtrip_semantics(track_session, session_key, path, "Report")["Report"]

    assert semantics["page_breaks"]["rows_count"] == 1
    assert semantics["page_breaks"]["rows_manualBreakCount"] == 0
    assert semantics["page_breaks"]["rows"] == [
        {"id": 8, "min": 0, "max": 16383, "man": False},
    ]
    assert semantics["page_breaks"]["columns_count"] == 1
    assert semantics["protected_ranges"] == [
        {"name": "Editable", "sqref": "B4:C6", "password": "ABCD"},
    ]
    view = semantics["sheet_views"][0]
    assert view["topLeftCell"] == "C4"
    assert view["pane"]["topLeftCell"] == "C5"
    assert view["selections"] == [{"activeCell": "D4", "sqref": "D4:E6"}]
    assert "$A$2:$B$5" in semantics["print_area"]
    assert "$D$4:$E$8" in semantics["print_area"]
    assert semantics["print_titles"] == {"rows": "$1:$2", "cols": None}

    root = _worksheet_root(path)
    row_breaks = root.find("main:rowBreaks", NS)
    assert row_breaks is not None
    assert row_breaks.attrib == {"count": "1", "manualBreakCount": "0"}
    assert [item.attrib["id"] for item in row_breaks.findall("main:brk", NS)] == ["8"]
    protected = root.find("main:protectedRanges/main:protectedRange", NS)
    assert protected is not None and protected.attrib["sqref"] == "B4:C6"

    defined_names = _workbook_root(path).findall("main:definedNames/main:definedName", NS)
    by_name = {item.attrib["name"]: item for item in defined_names}
    assert "$A$2:$B$5" in (by_name["_xlnm.Print_Area"].text or "")
    assert "$D$4:$E$8" in (by_name["_xlnm.Print_Area"].text or "")
