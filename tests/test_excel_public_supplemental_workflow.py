"""Public-tool creation and preservation coverage for the 28 supplemental cases."""

from __future__ import annotations

import base64
import json
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers" / "excel"))

import main as M  # noqa: E402


BASE_FIXTURE = Path(r"D:\data-test\excel-preservation-fixtures\sources\00-base.xlsx")
LEGACY_REPORT = Path(
    r"D:\data-test\excel-preservation-fixtures\reports\supplemental-coverage-latest.json"
)

SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
MAIN = f"{{{SHEET_NS}}}"
REL_ID = f"{{{REL_NS}}}id"

SUPPLEMENTAL_CASES = (
    "table.attributes_lost",
    "table.filter_criteria_lost",
    "table.totals_metadata_lost",
    "table.formulas_lost",
    "table.style_flags_present",
    "printing.print_area_removed",
    "printing.print_titles_source_available",
    "rich_text.five_runs_flattened",
    "rich_text.phonetic_lost",
    "rich_text.whitespace_fixture_present",
    "rich_text.charset_scheme_runs_flattened",
    "book1.b3_five_runs_flattened",
    "book1.blank_cells_added",
    "advanced.source.vba",
    "advanced.source.signatures",
    "advanced.source.vml",
    "advanced.source.drawing_chart_image",
    "advanced.source.printer_settings",
    "advanced.source.pivot_table",
    "advanced.source.pivot_cache",
    "advanced.source.slicer",
    "advanced.source.timeline",
    "advanced.source.external_connections_calcchain",
    "advanced.source.activex_ole",
    "advanced.source.threaded_comments_persons",
    "advanced.source.custom_ui_xml_model",
    "advanced.source.custom_properties",
    "advanced.current_part_loss_detectable",
)

ADVANCED_PART_GROUPS = {
    "advanced.source.vba": {"xl/vbaProject.bin"},
    "advanced.source.signatures": {"_xmlsignatures/origin.sigs", "_xmlsignatures/sig1.xml"},
    "advanced.source.vml": {"xl/vmlDrawings/vmlDrawing1.vml"},
    "advanced.source.drawing_chart_image": {
        "xl/charts/chart1.xml", "xl/drawings/drawing1.xml", "xl/media/image1.png",
    },
    "advanced.source.printer_settings": {"xl/printerSettings/printerSettings1.bin"},
    "advanced.source.pivot_table": {"xl/pivotTables/pivotTable1.xml"},
    "advanced.source.pivot_cache": {
        "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml",
    },
    "advanced.source.slicer": {"xl/slicers/slicer1.xml"},
    "advanced.source.timeline": {"xl/timelines/timeline1.xml"},
    "advanced.source.external_connections_calcchain": {
        "xl/calcChain.xml", "xl/connections.xml", "xl/externalLinks/externalLink1.xml",
    },
    "advanced.source.activex_ole": {
        "xl/activeX/activeX1.bin", "xl/activeX/activeX1.xml", "xl/embeddings/oleObject1.bin",
    },
    "advanced.source.threaded_comments_persons": {
        "xl/persons/person.xml", "xl/threadedComments/threadedComment1.xml",
    },
    "advanced.source.custom_ui_xml_model": {
        "customUI/customUI.xml", "customXml/item1.xml", "customXml/itemProps1.xml", "xl/model/model.xml",
    },
    "advanced.source.custom_properties": {"docProps/custom.xml"},
}


def _load_key(load_result: str) -> str:
    return load_result.split("session_key='")[1].split("'")[0]


def _relationship_part(source_part: str) -> str:
    if source_part == "/":
        return "_rels/.rels"
    directory = posixpath.dirname(source_part)
    filename = posixpath.basename(source_part)
    prefix = f"{directory}/" if directory else ""
    return f"{prefix}_rels/{filename}.rels"


def _package_relationships(path: Path, source_part: str) -> list[dict]:
    rel_path = _relationship_part(source_part)
    with zipfile.ZipFile(path, "r") as archive:
        if rel_path not in archive.namelist():
            return []
        root = ET.fromstring(archive.read(rel_path))
    return [
        {
            "id": item.attrib["Id"],
            "type": item.attrib["Type"],
            "target": item.attrib["Target"],
            "target_mode": item.attrib.get("TargetMode", "Internal"),
        }
        for item in root
    ]


def _relationship(
    relationship_id: str,
    relationship_type: str,
    target: str,
    target_mode: str = "Internal",
) -> dict:
    return {
        "id": relationship_id,
        "type": relationship_type,
        "target": target,
        "target_mode": target_mode,
    }


def _rich_runs() -> list[dict]:
    return [
        {"text": "Bold ", "font": {"bold": True}},
        {
            "text": "red ",
            "font": {"color": {"type": "rgb", "rgb": "FFFF0000"}},
        },
        {"text": "italic ", "font": {"italic": True}},
        {
            "text": "strike\neffects ",
            "font": {"strike": True, "underline": "single"},
        },
        {"text": "plain", "font": {}},
    ]


def _create_semantic_stage(stage_path: Path) -> None:
    session_key = _load_key(M.excel_load(str(BASE_FIXTURE)))
    try:
        M.excel_edit_cells(
            session_key,
            "Scores",
            [
                {"row_index": 0, "edits": {0: "Name", 1: "Score"}},
                {"row_index": 1, "edits": {0: "Alice", 1: 10}},
                {"row_index": 2, "edits": {0: "Total", 1: 10}},
            ],
        )
        M.excel_add_table(
            session_key,
            "Scores",
            "SupplementalTable",
            "A1:B3",
            table={
                "tableType": "worksheet",
                "headerRowCount": 1,
                "totalsRowCount": 1,
                "totalsRowShown": True,
                "insertRow": True,
                "insertRowShift": True,
                "published": True,
                "columns": [
                    {"id": 1, "name": "Name", "totalsRowLabel": "Total"},
                    {
                        "id": 2,
                        "name": "Score",
                        "totalsRowFunction": "sum",
                        "calculatedColumnFormula": "[@Score]*2",
                        "totalsRowFormula": "SUBTOTAL(109,[Score])",
                    },
                ],
                "auto_filter": {
                    "ref": "A1:B3",
                    "filter_columns": [{"colId": 0, "filters": ["Alice"]}],
                },
            },
            style={
                "name": "TableStyleMedium2",
                "showFirstColumn": True,
                "showLastColumn": True,
                "showRowStripes": False,
                "showColumnStripes": True,
            },
        )
        M.excel_set_print_area(session_key, "Scores", "$A$1:$B$3")
        M.excel_set_print_titles(session_key, "Scores", repeated_rows="$1:$1")
        M.excel_edit_rich_text(
            session_key,
            "Scores",
            "C1",
            operations=[{"op": "replace_runs", "runs": _rich_runs()}],
        )
        M.excel_edit_rich_text(
            session_key,
            "Scores",
            "C1",
            operations=[
                {
                    "op": "set_phonetic",
                    "runs": [{"text": "hint", "start": 0, "end": 4}],
                    "properties": {
                        "type": "fullwidthKatakana",
                        "alignment": "left",
                    },
                }
            ],
        )
        M.excel_edit_rich_text(
            session_key,
            "Scores",
            "C2",
            operations=[
                {
                    "op": "replace_runs",
                    "runs": [{"text": "  leading\ntrailing  ", "font": {}}],
                }
            ],
        )
        M.excel_edit_rich_text(
            session_key,
            "Scores",
            "C3",
            operations=[
                {
                    "op": "replace_runs",
                    "runs": [
                        {
                            "text": "Charset",
                            "font": {"name": "Arial", "charset": 204, "family": 2},
                        },
                        {
                            "text": " Scheme",
                            "font": {"name": "Calibri", "scheme": "minor"},
                        },
                    ],
                }
            ],
        )
        M.excel_add_sheet(session_key, "Book1")
        M.excel_edit_rich_text(
            session_key,
            "Book1",
            "B3",
            operations=[{"op": "replace_runs", "runs": _rich_runs()}],
        )
        M.excel_set_document_properties(
            session_key,
            custom=[
                {
                    "name": "Reviewer",
                    "type": "StringProperty",
                    "value": "Public supplemental workflow",
                },
                {"name": "RevisionNum", "type": "IntProperty", "value": 28},
                {"name": "Approved", "type": "BoolProperty", "value": True},
            ],
        )
        M.excel_save_as_copy(session_key, str(stage_path))
    finally:
        M.excel_close(session_key)


def _workbook_with_hidden_print_titles(workbook_xml: str) -> str:
    ET.register_namespace("", SHEET_NS)
    ET.register_namespace("r", REL_NS)
    root = ET.fromstring(workbook_xml)
    target = next(
        item
        for item in root.findall(f"{MAIN}definedNames/{MAIN}definedName")
        if item.attrib.get("name") == "_xlnm.Print_Titles"
        and item.attrib.get("localSheetId") == "0"
    )
    target.set("hidden", "1")
    pivot_caches = root.find(f"{MAIN}pivotCaches")
    if pivot_caches is None:
        pivot_caches = ET.SubElement(root, f"{MAIN}pivotCaches")
    pivot_cache = ET.SubElement(pivot_caches, f"{MAIN}pivotCache")
    pivot_cache.set("cacheId", "1")
    pivot_cache.set(REL_ID, "rIdSupplementalPivotCache")
    return ET.tostring(root, encoding="unicode")


def _advanced_upserts(workbook_xml: str) -> list[dict]:
    png_1x1 = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
    )
    xml_parts = {
        "xl/workbook.xml": (
            workbook_xml,
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        ),
        "_xmlsignatures/sig1.xml": (
            '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#">'
            "<SignedInfo/><SignatureValue>U1VQUExFTUVOVEFM</SignatureValue></Signature>",
            "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
        ),
        "xl/vmlDrawings/vmlDrawing1.vml": (
            '<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="SupplementalVML"/></xml>',
            "application/vnd.openxmlformats-officedocument.vmlDrawing",
        ),
        "xl/drawings/drawing1.xml": (
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>',
            "application/vnd.openxmlformats-officedocument.drawing+xml",
        ),
        "xl/charts/chart1.xml": (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            "<c:chart/></c:chartSpace>",
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        ),
        "xl/pivotTables/pivotTable1.xml": (
            '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'name="SupplementalPivot" cacheId="1" dataCaption="Values">'
            '<location ref="A1:B3" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'
            '</pivotTableDefinition>',
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
        ),
        "xl/pivotCache/pivotCacheDefinition1.xml": (
            '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'recordCount="0"><cacheSource type="worksheet"/></pivotCacheDefinition>',
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
        ),
        "xl/pivotCache/pivotCacheRecords1.xml": (
            '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>',
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
        ),
        "xl/slicers/slicer1.xml": (
            '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>',
            "application/vnd.ms-excel.slicer+xml",
        ),
        "xl/timelines/timeline1.xml": (
            '<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"/>',
            "application/vnd.ms-excel.timeline+xml",
        ),
        "xl/connections.xml": (
            '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>',
            "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
        ),
        "xl/externalLinks/externalLink1.xml": (
            '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<externalBook/></externalLink>",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
        ),
        "xl/calcChain.xml": (
            '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<c r="B2" i="1"/></calcChain>',
            "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
        ),
        "xl/activeX/activeX1.xml": (
            '<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" classid="Supplemental"/>',
            "application/vnd.ms-office.activeX+xml",
        ),
        "xl/threadedComments/threadedComment1.xml": (
            '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">'
            '<threadedComment ref="C1" personId="{00000000-0000-0000-0000-000000000001}" '
            'id="{00000000-0000-0000-0000-000000000002}" dT="2026-08-11T00:00:00Z">'
            "Supplemental</threadedComment></ThreadedComments>",
            "application/vnd.ms-excel.threadedcomments+xml",
        ),
        "xl/persons/person.xml": (
            '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/person">'
            '<person displayName="DocLoupe" id="{00000000-0000-0000-0000-000000000001}" '
            'userId="docloupe" providerId="None"/></personList>',
            "application/vnd.ms-excel.person+xml",
        ),
        "customUI/customUI.xml": (
            '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>',
            "application/vnd.ms-office.customUI+xml",
        ),
        "customXml/item1.xml": (
            '<supplemental xmlns="urn:docloupe:supplemental">public package data</supplemental>',
            "application/xml",
        ),
        "customXml/itemProps1.xml": (
            '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" '
            'ds:itemID="{00000000-0000-0000-0000-000000000028}"/>',
            "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
        ),
        "xl/model/model.xml": (
            '<model xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/model"/>',
            "application/vnd.ms-excel.model+xml",
        ),
    }
    binary_parts = {
        "xl/vbaProject.bin": (
            b"DOCLOUPE-SUPPLEMENTAL-VBA",
            "application/vnd.ms-office.vbaProject",
        ),
        "_xmlsignatures/origin.sigs": (
            b"",
            "application/vnd.openxmlformats-package.digital-signature-origin",
        ),
        "xl/media/image1.png": (png_1x1, "image/png"),
        "xl/printerSettings/printerSettings1.bin": (
            b"DOCLOUPE-PRINTER-SETTINGS",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
        ),
        "xl/activeX/activeX1.bin": (
            b"DOCLOUPE-ACTIVEX",
            "application/vnd.ms-office.activeX",
        ),
        "xl/embeddings/oleObject1.bin": (
            b"DOCLOUPE-OLE",
            "application/vnd.openxmlformats-officedocument.oleObject",
        ),
    }
    upserts = [
        {"path": path, "content": content, "content_type": content_type}
        for path, (content, content_type) in xml_parts.items()
    ]
    upserts.extend(
        {
            "path": path,
            "content": base64.b64encode(content).decode("ascii"),
            "encoding": "base64",
            "content_type": content_type,
        }
        for path, (content, content_type) in binary_parts.items()
    )
    return upserts


def _advanced_relationships(stage_path: Path) -> list[dict]:
    office_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_signature_rel = (
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature"
    )
    root_relationships = _package_relationships(stage_path, "/") + [
        _relationship(
            "rIdSupplementalSignature",
            f"{package_signature_rel}/origin",
            "_xmlsignatures/origin.sigs",
        ),
        _relationship(
            "rIdSupplementalCustomUI",
            "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
            "customUI/customUI.xml",
        ),
        _relationship(
            "rIdSupplementalCustomXml",
            f"{office_rel}/customXml",
            "customXml/item1.xml",
        ),
    ]
    workbook_relationships = _package_relationships(stage_path, "xl/workbook.xml") + [
        _relationship("rIdSupplementalVba", f"{office_rel}/vbaProject", "vbaProject.bin"),
        _relationship(
            "rIdSupplementalConnections",
            f"{office_rel}/connections",
            "connections.xml",
        ),
        _relationship(
            "rIdSupplementalExternalLink",
            f"{office_rel}/externalLink",
            "externalLinks/externalLink1.xml",
        ),
        _relationship(
            "rIdSupplementalCalcChain",
            f"{office_rel}/calcChain",
            "calcChain.xml",
        ),
        _relationship(
            "rIdSupplementalPivotCache",
            f"{office_rel}/pivotCacheDefinition",
            "pivotCache/pivotCacheDefinition1.xml",
        ),
        _relationship(
            "rIdSupplementalPersons",
            "http://schemas.microsoft.com/office/2017/10/relationships/person",
            "persons/person.xml",
        ),
        _relationship(
            "rIdSupplementalModel",
            "http://schemas.microsoft.com/office/2017/06/relationships/model",
            "model/model.xml",
        ),
    ]
    worksheet_relationships = _package_relationships(
        stage_path, "xl/worksheets/sheet1.xml"
    ) + [
        _relationship(
            "rIdSupplementalDrawing",
            f"{office_rel}/drawing",
            "../drawings/drawing1.xml",
        ),
        _relationship(
            "rIdSupplementalVml",
            f"{office_rel}/vmlDrawing",
            "../vmlDrawings/vmlDrawing1.vml",
        ),
        _relationship(
            "rIdSupplementalPrinter",
            f"{office_rel}/printerSettings",
            "../printerSettings/printerSettings1.bin",
        ),
        _relationship(
            "rIdSupplementalPivotTable",
            f"{office_rel}/pivotTable",
            "../pivotTables/pivotTable1.xml",
        ),
        _relationship(
            "rIdSupplementalThreadedComment",
            "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
            "../threadedComments/threadedComment1.xml",
        ),
        _relationship(
            "rIdSupplementalActiveX",
            f"{office_rel}/control",
            "../activeX/activeX1.xml",
        ),
        _relationship(
            "rIdSupplementalOle",
            f"{office_rel}/oleObject",
            "../embeddings/oleObject1.bin",
        ),
        _relationship(
            "rIdSupplementalSlicer",
            "http://schemas.microsoft.com/office/2007/relationships/slicer",
            "../slicers/slicer1.xml",
        ),
        _relationship(
            "rIdSupplementalTimeline",
            "http://schemas.microsoft.com/office/2011/relationships/timeline",
            "../timelines/timeline1.xml",
        ),
    ]
    return [
        {"source_part": "/", "relationships": root_relationships},
        {"source_part": "xl/workbook.xml", "relationships": workbook_relationships},
        {
            "source_part": "xl/worksheets/sheet1.xml",
            "relationships": worksheet_relationships,
        },
        {
            "source_part": "_xmlsignatures/origin.sigs",
            "relationships": [
                _relationship(
                    "rIdSignature1",
                    f"{package_signature_rel}/signature",
                    "sig1.xml",
                )
            ],
        },
        {
            "source_part": "xl/drawings/drawing1.xml",
            "relationships": [
                _relationship("rIdChart1", f"{office_rel}/chart", "../charts/chart1.xml"),
                _relationship("rIdImage1", f"{office_rel}/image", "../media/image1.png"),
            ],
        },
        {
            "source_part": "xl/pivotTables/pivotTable1.xml",
            "relationships": [
                _relationship(
                    "rIdPivotCache1",
                    f"{office_rel}/pivotCacheDefinition",
                    "../pivotCache/pivotCacheDefinition1.xml",
                )
            ],
        },
        {
            "source_part": "xl/pivotCache/pivotCacheDefinition1.xml",
            "relationships": [
                _relationship(
                    "rIdPivotRecords1",
                    f"{office_rel}/pivotCacheRecords",
                    "pivotCacheRecords1.xml",
                )
            ],
        },
        {
            "source_part": "xl/activeX/activeX1.xml",
            "relationships": [
                _relationship(
                    "rIdActiveXBinary",
                    f"{office_rel}/activeXControlBinary",
                    "activeX1.bin",
                )
            ],
        },
        {
            "source_part": "customXml/item1.xml",
            "relationships": [
                _relationship(
                    "rIdCustomXmlProps",
                    f"{office_rel}/customXmlProps",
                    "itemProps1.xml",
                )
            ],
        },
        {
            "source_part": "xl/externalLinks/externalLink1.xml",
            "relationships": [
                _relationship(
                    "rIdExternalBook",
                    f"{office_rel}/externalLinkPath",
                    "https://example.test/supplemental.xlsx",
                    "External",
                )
            ],
        },
    ]


def _create_public_supplemental_workbook(stage_path: Path, output_path: Path) -> None:
    _create_semantic_stage(stage_path)
    session_key = _load_key(M.excel_load(str(stage_path)))
    try:
        workbook_part = json.loads(
            M._package_tools["excel_read_package_part"](
                session_key,
                "xl/workbook.xml",
                output_mode="xml",
                max_bytes=262_144,
            )
        )
        workbook_xml = _workbook_with_hidden_print_titles(workbook_part["content"])
        M._package_tools["excel_apply_package_transaction"](
            session_key,
            upsert=_advanced_upserts(workbook_xml),
            relationships=_advanced_relationships(stage_path),
            max_summary_items=200,
        )
        M.excel_save_as_copy(session_key, str(output_path))
    finally:
        M.excel_close(session_key)


def _sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }
    result = {}
    for item in workbook.findall(f"{MAIN}sheets/{MAIN}sheet"):
        target = targets[item.attrib[REL_ID]]
        result[item.attrib["name"]] = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join("xl", target))
        )
    return result


def _rich_cell(archive: zipfile.ZipFile, sheet_part: str, cell_ref: str) -> dict:
    worksheet = ET.fromstring(archive.read(sheet_part))
    cell = next(
        item
        for item in worksheet.findall(f"{MAIN}sheetData/{MAIN}row/{MAIN}c")
        if item.attrib.get("r") == cell_ref
    )
    storage = cell.attrib.get("t")
    if storage == "s":
        shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        index = int(cell.findtext(f"{MAIN}v"))
        container = shared.findall(f"{MAIN}si")[index]
    elif storage == "inlineStr":
        container = cell.find(f"{MAIN}is")
    else:
        container = cell
    runs = container.findall(f"{MAIN}r") if container is not None else []
    text_nodes = []
    if container is not None:
        direct_text = container.find(f"{MAIN}t")
        if direct_text is not None:
            text_nodes.append(direct_text)
        text_nodes.extend(run.find(f"{MAIN}t") for run in runs)
    text_nodes = [item for item in text_nodes if item is not None]
    properties = []
    for run in runs:
        run_properties = run.find(f"{MAIN}rPr")
        properties.append(
            {
                child.tag.rsplit("}", 1)[-1]: child.attrib.get("val")
                for child in (list(run_properties) if run_properties is not None else [])
            }
        )
    return {
        "storage": storage,
        "text": "".join(item.text or "" for item in text_nodes),
        "run_count": len(runs),
        "run_properties": properties,
        "phonetic_run_count": len(container.findall(f"{MAIN}rPh"))
        if container is not None
        else 0,
        "space_preserved": any(
            item.attrib.get(f"{{{XML_NS}}}space") == "preserve" for item in text_nodes
        ),
    }


def _cell_refs(archive: zipfile.ZipFile, sheet_part: str) -> set[str]:
    worksheet = ET.fromstring(archive.read(sheet_part))
    return {
        item.attrib["r"]
        for item in worksheet.findall(f"{MAIN}sheetData/{MAIN}row/{MAIN}c")
    }


def _supplemental_status(path: Path, damaged_path: Path | None = None) -> dict[str, bool]:
    with zipfile.ZipFile(path, "r") as archive:
        parts = set(archive.namelist())
        sheet_parts = _sheet_parts(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        table = ET.fromstring(archive.read("xl/tables/table1.xml"))
        scores_rich = {
            cell: _rich_cell(archive, sheet_parts["Scores"], cell)
            for cell in ("C1", "C2", "C3")
        }
        book1_rich = _rich_cell(archive, sheet_parts["Book1"], "B3")
        book1_refs = _cell_refs(archive, sheet_parts["Book1"])

    table_columns = table.findall(f"{MAIN}tableColumns/{MAIN}tableColumn")
    score_column = next(item for item in table_columns if item.attrib.get("name") == "Score")
    name_column = next(item for item in table_columns if item.attrib.get("name") == "Name")
    table_style = table.find(f"{MAIN}tableStyleInfo")
    filter_values = table.findall(
        f"{MAIN}autoFilter/{MAIN}filterColumn/{MAIN}filters/{MAIN}filter"
    )
    defined_names = {
        item.attrib["name"]: item
        for item in workbook.findall(f"{MAIN}definedNames/{MAIN}definedName")
        if item.attrib.get("localSheetId") == "0"
    }
    print_area = defined_names.get("_xlnm.Print_Area")
    print_titles = defined_names.get("_xlnm.Print_Titles")

    statuses = {
        "table.attributes_lost": {
            key: table.attrib.get(key)
            for key in (
                "tableType",
                "totalsRowCount",
                "totalsRowShown",
                "insertRow",
                "insertRowShift",
                "published",
            )
        }
        == {
            "tableType": "worksheet",
            "totalsRowCount": "1",
            "totalsRowShown": "1",
            "insertRow": "1",
            "insertRowShift": "1",
            "published": "1",
        },
        "table.filter_criteria_lost": any(
            item.attrib.get("val") == "Alice" for item in filter_values
        ),
        "table.totals_metadata_lost": (
            name_column.attrib.get("totalsRowLabel") == "Total"
            and score_column.attrib.get("totalsRowFunction") == "sum"
        ),
        "table.formulas_lost": (
            score_column.findtext(f"{MAIN}calculatedColumnFormula") == "[@Score]*2"
            and score_column.findtext(f"{MAIN}totalsRowFormula")
            == "SUBTOTAL(109,[Score])"
        ),
        "table.style_flags_present": table_style is not None
        and table_style.attrib
        == {
            "name": "TableStyleMedium2",
            "showFirstColumn": "1",
            "showLastColumn": "1",
            "showRowStripes": "0",
            "showColumnStripes": "1",
        },
        "printing.print_area_removed": print_area is not None
        and print_area.attrib.get("localSheetId") == "0"
        and "$A$1:$B$3" in (print_area.text or ""),
        "printing.print_titles_source_available": print_titles is not None
        and print_titles.attrib.get("localSheetId") == "0"
        and print_titles.attrib.get("hidden") == "1"
        and "$1:$1" in (print_titles.text or ""),
        "rich_text.five_runs_flattened": scores_rich["C1"]["run_count"] == 5,
        "rich_text.phonetic_lost": scores_rich["C1"]["phonetic_run_count"] == 1,
        "rich_text.whitespace_fixture_present": (
            scores_rich["C2"]["space_preserved"]
            and scores_rich["C2"]["text"] == "  leading\ntrailing  "
        ),
        "rich_text.charset_scheme_runs_flattened": (
            scores_rich["C3"]["run_count"] == 2
            and any(
                item.get("charset") == "204"
                for item in scores_rich["C3"]["run_properties"]
            )
            and any(
                item.get("scheme") == "minor"
                for item in scores_rich["C3"]["run_properties"]
            )
        ),
        "book1.b3_five_runs_flattened": book1_rich["run_count"] == 5,
        "book1.blank_cells_added": book1_refs == {"B3"},
    }
    statuses.update(
        {name: required.issubset(parts) for name, required in ADVANCED_PART_GROUPS.items()}
    )

    loss_detected = False
    if damaged_path is not None:
        report = json.loads(
            M.excel_verify_preservation(
                after_path=str(damaged_path),
                before_path=str(path),
                fixture_id="public-supplemental-part-loss",
                max_differences=5000,
            )
        )
        loss_detected = (
            report["preservation_ok"] is False
            and report["classification_counts"]["UNAPPROVED_LOSS"] > 0
            and "xl/calcChain.xml" in report["part_diff"]["removed"]
        )
    statuses["advanced.current_part_loss_detectable"] = loss_detected
    assert set(statuses) == set(SUPPLEMENTAL_CASES)
    return statuses


def _supplemental_report(path: Path, damaged_path: Path) -> dict:
    statuses = _supplemental_status(path, damaged_path)
    failed = [name for name in SUPPLEMENTAL_CASES if not statuses[name]]
    return {
        "path": str(path),
        "checked_case_count": len(SUPPLEMENTAL_CASES),
        "preserved_case_count": len(SUPPLEMENTAL_CASES) - len(failed),
        "failed_case_count": len(failed),
        "failed_cases": failed,
    }


def _create_damaged_copy(source_path: Path, damaged_path: Path) -> None:
    session_key = _load_key(M.excel_load(str(source_path)))
    try:
        relationships = [
            item
            for item in _package_relationships(source_path, "xl/workbook.xml")
            if not (
                item["type"].endswith("/calcChain")
                or item["target"].endswith("calcChain.xml")
            )
        ]
        M._package_tools["excel_set_package_relationships"](
            session_key,
            "xl/workbook.xml",
            relationships,
        )
        M._package_tools["excel_delete_package_part"](
            session_key,
            "xl/calcChain.xml",
        )
        M.excel_save_as_copy(session_key, str(damaged_path))
    finally:
        M.excel_close(session_key)


def test_public_tools_create_reload_and_preserve_supplemental_28(tmp_path, monkeypatch):
    if not BASE_FIXTURE.is_file():
        pytest.skip(f"Immutable base fixture is unavailable: {BASE_FIXTURE}")
    monkeypatch.setenv("DOCLOUPE_EXCEL_BACKUP_DIR", str(tmp_path / "backups"))
    stage_path = tmp_path / "public-supplemental-stage.xlsx"
    output_path = tmp_path / "public-supplemental.xlsm"
    roundtrip_path = tmp_path / "public-supplemental-roundtrip.xlsm"
    damaged_path = tmp_path / "public-supplemental-damaged.xlsm"

    _create_public_supplemental_workbook(stage_path, output_path)
    assert json.loads(M.excel_validate_workbook(str(output_path)))["valid"] is True

    _create_damaged_copy(output_path, damaged_path)
    assert json.loads(M.excel_validate_workbook(str(damaged_path)))["valid"] is True
    created_report = _supplemental_report(output_path, damaged_path)
    assert created_report == {
        "path": str(output_path),
        "checked_case_count": 28,
        "preserved_case_count": 28,
        "failed_case_count": 0,
        "failed_cases": [],
    }

    session_key = _load_key(M.excel_load(str(output_path)))
    try:
        M.excel_save_as_copy(session_key, str(roundtrip_path))
    finally:
        M.excel_close(session_key)

    assert json.loads(M.excel_validate_workbook(str(roundtrip_path)))["valid"] is True
    roundtrip_report = _supplemental_report(roundtrip_path, damaged_path)
    assert roundtrip_report["checked_case_count"] == 28
    assert roundtrip_report["preserved_case_count"] == 28
    assert roundtrip_report["failed_case_count"] == 0
    preservation = json.loads(
        M.excel_verify_preservation(
            after_path=str(roundtrip_path),
            before_path=str(output_path),
            fixture_id="public-supplemental-roundtrip",
            max_differences=5000,
        )
    )
    assert preservation["preservation_ok"] is True
    assert preservation["unapproved_difference_count"] == 0


def test_add_table_handles_loaded_workbook_with_none_table_collection(tmp_path):
    source_path = tmp_path / "plain-table-source.xlsx"
    output_path = tmp_path / "plain-table-output.xlsx"
    created = json.loads(
        M.excel_create_workbook(sheet_names=["S"], target_path=str(source_path))
    )
    session_key = created["session_key"]
    try:
        M.excel_edit_cells(session_key, "S", [
            {"row_index": 0, "edits": {0: "Name", 1: "Value"}},
            {"row_index": 1, "edits": {0: "A", 1: 1}},
        ])
        M.excel_save(session_key)
    finally:
        M.excel_close(session_key)

    loaded_key = _load_key(M.excel_load(str(source_path)))
    try:
        M.excel_add_table(loaded_key, "S", "LoadedTable", "A1:B2")
        M.excel_save_as_copy(loaded_key, str(output_path))
    finally:
        M.excel_close(loaded_key)

    assert json.loads(M.excel_validate_workbook(str(output_path)))["valid"] is True
    with zipfile.ZipFile(output_path, "r") as archive:
        assert "xl/tables/table1.xml" in archive.namelist()


def test_phonetic_creation_defaults_required_font_id_and_reloads(tmp_path):
    output_path = tmp_path / "phonetic-default.xlsx"
    created = json.loads(
        M.excel_create_workbook(sheet_names=["S"], target_path=str(output_path))
    )
    session_key = created["session_key"]
    try:
        M.excel_edit_rich_text(
            session_key,
            "S",
            "A1",
            operations=[{"op": "replace_runs", "runs": _rich_runs()}],
        )
        M.excel_edit_rich_text(
            session_key,
            "S",
            "A1",
            operations=[{
                "op": "set_phonetic",
                "runs": [{"text": "hint", "start": 0, "end": 4}],
                "properties": {"type": "fullwidthKatakana"},
            }],
        )
        M.excel_save(session_key)
    finally:
        M.excel_close(session_key)

    loaded_key = _load_key(M.excel_load(str(output_path)))
    try:
        model = json.loads(M.excel_get_rich_text(loaded_key, "S", "A1"))
        assert model["phonetic_properties"]["fontId"] == "0"
        assert model["phonetic_runs"][0]["text"] == "hint"
    finally:
        M.excel_close(loaded_key)


def test_explicit_package_transaction_promotes_xlsx_to_xlsm_and_keeps_root_ids_unique(tmp_path):
    source_path = tmp_path / "promotion-source.xlsx"
    output_path = tmp_path / "promotion-output.xlsm"
    created = json.loads(
        M.excel_create_workbook(sheet_names=["S"], target_path=str(source_path))
    )
    session_key = created["session_key"]
    try:
        M.excel_edit_cells(
            session_key,
            "S",
            [{"row_index": 0, "edits": {0: "macro promotion"}}],
        )
        M.excel_save(session_key)
    finally:
        M.excel_close(session_key)

    loaded_key = _load_key(M.excel_load(str(source_path)))
    try:
        workbook_part = json.loads(
            M._package_tools["excel_read_package_part"](
                loaded_key,
                "xl/workbook.xml",
                output_mode="xml",
                max_bytes=262_144,
            )
        )
        root_relationships = _package_relationships(source_path, "/") + [
            _relationship(
                "rIdPromotionCustomXml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
                "customXml/promotion.xml",
            )
        ]
        M._package_tools["excel_apply_package_transaction"](
            loaded_key,
            upsert=[
                {
                    "path": "xl/workbook.xml",
                    "content": workbook_part["content"],
                    "content_type": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                },
                {
                    "path": "xl/vbaProject.bin",
                    "content": base64.b64encode(b"PROMOTION-VBA").decode("ascii"),
                    "encoding": "base64",
                    "content_type": "application/vnd.ms-office.vbaProject",
                },
                {
                    "path": "customXml/promotion.xml",
                    "content": "<promotion/>",
                    "content_type": "application/xml",
                },
            ],
            relationships=[
                {"source_part": "/", "relationships": root_relationships}
            ],
        )
        M.excel_save_as_copy(loaded_key, str(output_path))
    finally:
        M.excel_close(loaded_key)

    assert json.loads(M.excel_validate_workbook(str(output_path)))["valid"] is True
    with zipfile.ZipFile(output_path, "r") as archive:
        root = ET.fromstring(archive.read("_rels/.rels"))
        relationship_ids = [item.attrib["Id"] for item in root]
        assert len(relationship_ids) == len(set(relationship_ids))
        assert "xl/vbaProject.bin" in archive.namelist()
    promoted_key = _load_key(M.excel_load(str(output_path)))
    M.excel_close(promoted_key)


def test_supplemental_checker_matches_legacy_case_inventory_when_available():
    if not LEGACY_REPORT.is_file():
        pytest.skip(f"Legacy supplemental report is unavailable: {LEGACY_REPORT}")
    report = json.loads(LEGACY_REPORT.read_text(encoding="utf-8"))
    assert report["case_count"] == 28
    assert report["reproduced_count"] == 28
    assert tuple(item["name"] for item in report["results"]) == SUPPLEMENTAL_CASES
