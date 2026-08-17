"""Public acceptance coverage for DrawingML creation and preservation."""

import json
import posixpath
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile
import zlib
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers" / "excel"))

import main as M  # noqa: E402


CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAWING_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
OFFICE_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

DRAWING_REL_TYPE = f"{OFFICE_RELS_NS}/drawing"
IMAGE_REL_TYPE = f"{OFFICE_RELS_NS}/image"
CHART_REL_TYPE = f"{OFFICE_RELS_NS}/chart"
DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"
CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"


def _load_key(load_result: str) -> str:
    return load_result.split("session_key='")[1].split("'")[0]


def _png_chunk(chunk_type: bytes, payload: bytes) -> bytes:
    checksum = zlib.crc32(chunk_type)
    checksum = zlib.crc32(payload, checksum) & 0xFFFFFFFF
    return struct.pack(">I", len(payload)) + chunk_type + payload + struct.pack(">I", checksum)


def _write_tiny_png(path: Path, width: int = 2, height: int = 3) -> bytes:
    rows = []
    for row in range(height):
        pixels = bytearray()
        for column in range(width):
            pixels.extend((220, 40 + row * 20, 60 + column * 30))
        rows.append(b"\x00" + bytes(pixels))
    payload = (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + _png_chunk(b"IDAT", zlib.compress(b"".join(rows)))
        + _png_chunk(b"IEND", b"")
    )
    path.write_bytes(payload)
    return payload


def _create_workbook(tmp_path: Path, filename: str) -> tuple[str, Path]:
    output_path = tmp_path / filename
    created = json.loads(
        M.excel_create_workbook(
            sheet_names=["Drawing"],
            active_sheet="Drawing",
            target_path=str(output_path),
        )
    )
    assert created["sheets"] == ["Drawing"]
    assert created["active_sheet"] == "Drawing"
    assert created["capabilities"]["requires_explicit_save"] is True
    assert not output_path.exists()
    return created["session_key"], output_path


def _save_validate_close_reload(session_key: str, output_path: Path) -> tuple[str, dict]:
    M.excel_save(session_key)
    assert output_path.exists()
    report = json.loads(M.excel_validate_workbook(str(output_path)))
    assert report["valid"] is True, report["errors"]
    assert report["errors"] == []
    M.excel_close(session_key)
    reloaded_key = _load_key(M.excel_load(str(output_path)))
    return reloaded_key, report


def _relationships(archive: zipfile.ZipFile, relationship_part: str) -> dict[str, dict]:
    if relationship_part not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(relationship_part))
    result = {}
    for relationship in root.findall(f"{{{PACKAGE_RELS_NS}}}Relationship"):
        result[relationship.attrib["Id"]] = {
            "type": relationship.attrib["Type"],
            "target": relationship.attrib["Target"],
            "target_mode": relationship.attrib.get("TargetMode"),
        }
    return result


def _resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _drawing_relationship_part(drawing_part: str) -> str:
    folder, filename = drawing_part.rsplit("/", 1)
    return f"{folder}/_rels/{filename}.rels"


def _content_types(archive: zipfile.ZipFile) -> tuple[dict[str, str], dict[str, str]]:
    root = ET.fromstring(archive.read("[Content_Types].xml"))
    defaults = {
        item.attrib["Extension"].lower(): item.attrib["ContentType"]
        for item in root.findall(f"{{{CONTENT_TYPES_NS}}}Default")
    }
    overrides = {
        item.attrib["PartName"].lstrip("/"): item.attrib["ContentType"]
        for item in root.findall(f"{{{CONTENT_TYPES_NS}}}Override")
    }
    return defaults, overrides


def _drawing_objects(drawing_root: ET.Element) -> list[dict]:
    objects = []
    for anchor in list(drawing_root):
        anchor_type = anchor.tag.rsplit("}", 1)[-1]
        if not anchor_type.endswith("Anchor"):
            continue

        picture = anchor.find(f"{{{DRAWING_NS}}}pic")
        graphic_frame = anchor.find(f"{{{DRAWING_NS}}}graphicFrame")
        shape = anchor.find(f"{{{DRAWING_NS}}}sp")
        if picture is not None:
            kind = "picture"
            element = picture
            relationship_node = picture.find(f".//{{{DRAWING_MAIN_NS}}}blip")
            relationship_id = (
                relationship_node.attrib.get(f"{{{OFFICE_RELS_NS}}}embed")
                if relationship_node is not None
                else None
            )
        elif graphic_frame is not None and graphic_frame.find(f".//{{{CHART_NS}}}chart") is not None:
            kind = "chart"
            element = graphic_frame
            relationship_node = graphic_frame.find(f".//{{{CHART_NS}}}chart")
            relationship_id = relationship_node.attrib.get(f"{{{OFFICE_RELS_NS}}}id")
        elif shape is not None:
            kind = "shape"
            element = shape
            relationship_id = None
        else:
            continue

        properties = element.find(f".//{{{DRAWING_NS}}}cNvPr")
        origin = anchor.find(f"{{{DRAWING_NS}}}from")
        extent = anchor.find(f"{{{DRAWING_NS}}}ext")
        geometry = element.find(f".//{{{DRAWING_MAIN_NS}}}prstGeom")
        objects.append(
            {
                "kind": kind,
                "name": properties.attrib.get("name") if properties is not None else None,
                "relationship_id": relationship_id,
                "anchor_type": anchor_type,
                "column": int(origin.findtext(f"{{{DRAWING_NS}}}col")) if origin is not None else None,
                "row": int(origin.findtext(f"{{{DRAWING_NS}}}row")) if origin is not None else None,
                "cx": int(extent.attrib["cx"]) if extent is not None else None,
                "cy": int(extent.attrib["cy"]) if extent is not None else None,
                "text": "".join(node.text or "" for node in element.findall(f".//{{{DRAWING_MAIN_NS}}}t")),
                "geometry": geometry.attrib.get("prst") if geometry is not None else None,
            }
        )
    return objects


def _package_state(path: Path) -> dict:
    sheet_part = "xl/worksheets/sheet1.xml"
    sheet_relationship_part = "xl/worksheets/_rels/sheet1.xml.rels"
    with zipfile.ZipFile(path, "r") as archive:
        parts = set(archive.namelist())
        assert sheet_part in parts
        sheet_root = ET.fromstring(archive.read(sheet_part))
        drawing_reference = sheet_root.find(f"{{{SPREADSHEET_NS}}}drawing")
        assert drawing_reference is not None, "worksheet has no DrawingML <drawing> reference"

        relationship_id = drawing_reference.attrib[f"{{{OFFICE_RELS_NS}}}id"]
        sheet_relationships = _relationships(archive, sheet_relationship_part)
        assert relationship_id in sheet_relationships
        drawing_relationship = sheet_relationships[relationship_id]
        assert drawing_relationship["type"] == DRAWING_REL_TYPE

        drawing_part = _resolve_target(sheet_part, drawing_relationship["target"])
        assert drawing_part in parts
        drawing_root = ET.fromstring(archive.read(drawing_part))
        drawing_relationship_part = _drawing_relationship_part(drawing_part)
        drawing_relationships = _relationships(archive, drawing_relationship_part)
        defaults, overrides = _content_types(archive)

        related_parts = {}
        for rel_id, relationship in drawing_relationships.items():
            if relationship["target_mode"] == "External":
                continue
            related_part = _resolve_target(drawing_part, relationship["target"])
            assert related_part in parts
            related_parts[rel_id] = {
                **relationship,
                "part": related_part,
                "payload": archive.read(related_part),
            }

    return {
        "parts": parts,
        "drawing_part": drawing_part,
        "drawing_relationship_part": drawing_relationship_part,
        "sheet_relationships": sheet_relationships,
        "drawing_relationships": drawing_relationships,
        "related_parts": related_parts,
        "defaults": defaults,
        "overrides": overrides,
        "objects": _drawing_objects(drawing_root),
    }


def _find_object(state: dict, kind: str, name: str | None = None) -> dict:
    candidates = [item for item in state["objects"] if item["kind"] == kind]
    if name is not None:
        candidates = [item for item in candidates if item["name"] == name]
    assert len(candidates) == 1, state["objects"]
    return candidates[0]


def _assert_anchor_and_dimensions(
    drawing_object: dict,
    expected_column: int,
    expected_row: int,
    expected_ratio: float,
) -> None:
    assert drawing_object["anchor_type"] == "oneCellAnchor"
    assert drawing_object["column"] == expected_column
    assert drawing_object["row"] == expected_row
    assert drawing_object["cx"] > 0
    assert drawing_object["cy"] > 0
    assert drawing_object["cx"] / drawing_object["cy"] == pytest.approx(expected_ratio, rel=0.01)


def _inventory(session_key: str) -> list[dict]:
    return json.loads(M.excel_get_shapes(session_key, "Drawing"))["Drawing"]


def _assert_drawing_content_type(state: dict) -> None:
    assert state["overrides"][state["drawing_part"]] == DRAWING_CONTENT_TYPE


def _relationship_payload(state: dict, drawing_object: dict, relationship_type: str) -> tuple[str, bytes]:
    relationship_id = drawing_object["relationship_id"]
    assert relationship_id in state["related_parts"]
    related = state["related_parts"][relationship_id]
    assert related["type"] == relationship_type
    return related["part"], related["payload"]


def test_public_add_image_creates_package_and_survives_adding_shape(tmp_path):
    session_key, output_path = _create_workbook(tmp_path, "public-image.xlsx")
    image_path = tmp_path / "tiny.png"
    image_payload = _write_tiny_png(image_path)

    M.excel_add_image(
        session_key,
        "Drawing",
        anchor="B2",
        source_path=str(image_path),
        width=120,
        height=60,
        name="Public Tiny Image",
    )
    session_key, report = _save_validate_close_reload(session_key, output_path)
    assert report["features"]["drawings"] >= 1
    assert report["features"]["media"] >= 1

    first_state = _package_state(output_path)
    _assert_drawing_content_type(first_state)
    assert first_state["defaults"]["png"] == "image/png"
    picture = _find_object(first_state, "picture", "Public Tiny Image")
    _assert_anchor_and_dimensions(picture, expected_column=1, expected_row=1, expected_ratio=2.0)
    image_part, stored_image = _relationship_payload(first_state, picture, IMAGE_REL_TYPE)
    assert image_part.startswith("xl/media/")
    assert stored_image == image_payload
    assert any(item["type"] == "picture" and item["name"] == "Public Tiny Image" for item in _inventory(session_key))
    picture_signature = {key: picture[key] for key in ("name", "column", "row", "cx", "cy")}

    M.excel_add_shape(
        session_key,
        "Drawing",
        shape_type="rect",
        anchor="F2",
        text="Companion",
        width=90,
        height=45,
        name="Companion Shape",
    )
    session_key, _ = _save_validate_close_reload(session_key, output_path)
    second_state = _package_state(output_path)
    surviving_picture = _find_object(second_state, "picture", "Public Tiny Image")
    assert {key: surviving_picture[key] for key in picture_signature} == picture_signature
    _, surviving_image = _relationship_payload(second_state, surviving_picture, IMAGE_REL_TYPE)
    assert surviving_image == image_payload
    assert _find_object(second_state, "shape", "Companion Shape")["text"] == "Companion"
    M.excel_close(session_key)


def test_public_add_chart_creates_package_and_survives_adding_image(tmp_path):
    session_key, output_path = _create_workbook(tmp_path, "public-chart.xlsx")
    M.excel_edit_cells(
        session_key,
        "Drawing",
        [
            {"row_index": 0, "edits": {0: "Category", 1: "Value"}},
            {"row_index": 1, "edits": {0: "Alpha", 1: 2}},
            {"row_index": 2, "edits": {0: "Beta", 1: 5}},
        ],
    )
    M.excel_add_chart(
        session_key,
        "Drawing",
        chart_type="bar",
        source_range="A1:B3",
        anchor="E2",
        title="Public Chart",
        width=8,
        height=5,
    )
    session_key, report = _save_validate_close_reload(session_key, output_path)
    assert report["features"]["drawings"] >= 1
    assert report["features"]["charts"] >= 1

    first_state = _package_state(output_path)
    _assert_drawing_content_type(first_state)
    chart = _find_object(first_state, "chart")
    _assert_anchor_and_dimensions(chart, expected_column=4, expected_row=1, expected_ratio=1.6)
    chart_part, chart_payload = _relationship_payload(first_state, chart, CHART_REL_TYPE)
    assert first_state["overrides"][chart_part] == CHART_CONTENT_TYPE
    chart_root = ET.fromstring(chart_payload)
    chart_text = "".join(node.text or "" for node in chart_root.findall(f".//{{{DRAWING_MAIN_NS}}}t"))
    assert "Public Chart" in chart_text
    assert any(item["type"] == "chart" for item in _inventory(session_key))

    image_path = tmp_path / "chart-companion.png"
    image_payload = _write_tiny_png(image_path, width=3, height=2)
    M.excel_add_image(
        session_key,
        "Drawing",
        anchor="A8",
        source_path=str(image_path),
        width=90,
        height=60,
        name="Chart Companion Image",
    )
    session_key, _ = _save_validate_close_reload(session_key, output_path)
    second_state = _package_state(output_path)
    surviving_chart = _find_object(second_state, "chart")
    surviving_chart_part, surviving_chart_payload = _relationship_payload(second_state, surviving_chart, CHART_REL_TYPE)
    assert surviving_chart_part == chart_part
    assert surviving_chart_payload == chart_payload
    companion_picture = _find_object(second_state, "picture", "Chart Companion Image")
    _, surviving_image = _relationship_payload(second_state, companion_picture, IMAGE_REL_TYPE)
    assert surviving_image == image_payload
    M.excel_close(session_key)


def test_public_add_shape_creates_package_and_survives_adding_chart(tmp_path):
    session_key, output_path = _create_workbook(tmp_path, "public-shape.xlsx")
    M.excel_add_shape(
        session_key,
        "Drawing",
        shape_type="rect",
        anchor="C3",
        text="Public Shape Text",
        width=3,
        height=1.5,
        style={"fill_color": "FF4472C4", "outline_color": "FF203864"},
        name="Public Shape",
    )
    session_key, report = _save_validate_close_reload(session_key, output_path)
    assert report["features"]["drawings"] >= 1

    first_state = _package_state(output_path)
    _assert_drawing_content_type(first_state)
    shape = _find_object(first_state, "shape", "Public Shape")
    _assert_anchor_and_dimensions(shape, expected_column=2, expected_row=2, expected_ratio=2.0)
    assert shape["geometry"] == "rect"
    assert shape["text"] == "Public Shape Text"
    assert any(
        item["type"] == "shape" and item["name"] == "Public Shape" and item["text"] == "Public Shape Text"
        for item in _inventory(session_key)
    )
    shape_signature = {key: shape[key] for key in ("name", "column", "row", "cx", "cy", "text", "geometry")}

    M.excel_edit_cells(
        session_key,
        "Drawing",
        [
            {"row_index": 0, "edits": {0: "Label", 1: "Amount"}},
            {"row_index": 1, "edits": {0: "One", 1: 1}},
            {"row_index": 2, "edits": {0: "Two", 1: 2}},
        ],
    )
    M.excel_add_chart(
        session_key,
        "Drawing",
        chart_type="line",
        source_range="A1:B3",
        anchor="G3",
        title="Shape Companion Chart",
        width=7,
        height=4,
    )
    session_key, _ = _save_validate_close_reload(session_key, output_path)
    second_state = _package_state(output_path)
    surviving_shape = _find_object(second_state, "shape", "Public Shape")
    assert {key: surviving_shape[key] for key in shape_signature} == shape_signature
    companion_chart = _find_object(second_state, "chart")
    companion_chart_part, _ = _relationship_payload(second_state, companion_chart, CHART_REL_TYPE)
    assert second_state["overrides"][companion_chart_part] == CHART_CONTENT_TYPE
    M.excel_close(session_key)


def test_repeated_same_session_save_does_not_duplicate_drawing_creations(tmp_path):
    session_key, output_path = _create_workbook(tmp_path, "public-drawing-repeat-save.xlsx")
    image_path = tmp_path / "repeat-save.png"
    image_payload = _write_tiny_png(image_path, width=4, height=2)
    try:
        M.excel_edit_cells(
            session_key,
            "Drawing",
            [
                {"row_index": 0, "edits": {0: "Category", 1: "Value"}},
                {"row_index": 1, "edits": {0: "Alpha", 1: 2}},
                {"row_index": 2, "edits": {0: "Beta", 1: 5}},
            ],
        )
        M.excel_add_chart(
            session_key,
            "Drawing",
            chart_type="line",
            source_range="A1:B3",
            anchor="E2",
            title="Repeated Save Chart",
            width=8,
            height=5,
        )
        M.excel_add_image(
            session_key,
            "Drawing",
            anchor="B8",
            source_path=str(image_path),
            width=120,
            height=60,
            name="Repeated Save Image",
        )
        M.excel_add_shape(
            session_key,
            "Drawing",
            shape_type="rect",
            anchor="H8",
            text="Repeated Save Shape",
            width=4,
            height=2,
            name="Repeated Save Shape",
        )

        M.excel_save(session_key)
        first_state = _package_state(output_path)
        M.excel_save(session_key)
        second_state = _package_state(output_path)

        assert second_state["drawing_part"] == first_state["drawing_part"]
        assert second_state["objects"] == first_state["objects"]
        assert len(second_state["objects"]) == 3
        first_relationships = sorted(
            (item["type"], item["part"], item["payload"])
            for item in first_state["related_parts"].values()
        )
        second_relationships = sorted(
            (item["type"], item["part"], item["payload"])
            for item in second_state["related_parts"].values()
        )
        assert second_relationships == first_relationships
        picture = _find_object(second_state, "picture", "Repeated Save Image")
        _, stored_image = _relationship_payload(second_state, picture, IMAGE_REL_TYPE)
        assert stored_image == image_payload
    finally:
        M.excel_close(session_key)


def test_structural_edits_shift_queued_drawing_anchors_and_chart_source(tmp_path):
    session_key, output_path = _create_workbook(tmp_path, "public-drawing-structural.xlsx")
    image_path = tmp_path / "structural.png"
    _write_tiny_png(image_path, width=4, height=2)
    try:
        M.excel_edit_cells(
            session_key,
            "Drawing",
            [
                {"row_index": 0, "edits": {0: "Category", 1: "Value"}},
                {"row_index": 1, "edits": {0: "Alpha", 1: 2}},
                {"row_index": 2, "edits": {0: "Beta", 1: 5}},
            ],
        )
        M.excel_add_chart(
            session_key,
            "Drawing",
            chart_type="line",
            source_range="A1:B3",
            anchor="E2",
            title="Shifted Chart",
        )
        M.excel_add_image(
            session_key,
            "Drawing",
            anchor="B8",
            source_path=str(image_path),
            name="Shifted Image",
        )
        M.excel_add_shape(
            session_key,
            "Drawing",
            shape_type="rect",
            anchor="H8",
            text="Shifted Shape",
            name="Shifted Shape",
        )

        M.excel_copy_row(session_key, "Drawing", 0, -1)
        M.excel_insert_column(session_key, "Drawing", after_col_index=-1)
        session_key, _ = _save_validate_close_reload(session_key, output_path)

        state = _package_state(output_path)
        chart = _find_object(state, "chart")
        picture = _find_object(state, "picture", "Shifted Image")
        shape = _find_object(state, "shape", "Shifted Shape")
        _assert_anchor_and_dimensions(chart, expected_column=5, expected_row=2, expected_ratio=2.0)
        _assert_anchor_and_dimensions(picture, expected_column=2, expected_row=8, expected_ratio=2.0)
        _assert_anchor_and_dimensions(shape, expected_column=8, expected_row=8, expected_ratio=2.0)

        _, chart_payload = _relationship_payload(state, chart, CHART_REL_TYPE)
        chart_xml = chart_payload.decode("utf-8")
        assert "'Drawing'!C2" in chart_xml
        assert "'Drawing'!$B$3:$B$4" in chart_xml
        assert "'Drawing'!$C$3:$C$4" in chart_xml
    finally:
        M.excel_close(session_key)


def test_public_shape_rich_text_creation_and_update_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCLOUPE_EXCEL_BACKUP_DIR", str(tmp_path / "backups"))
    session_key, output_path = _create_workbook(tmp_path, "public-shape-rich-text.xlsx")
    M.excel_add_shape(
        session_key,
        "Drawing",
        shape_type="rect",
        anchor="C3",
        rich_text={
            "runs": [
                {
                    "text": "Red",
                    "font": {
                        "bold": True,
                        "size": 14,
                        "name": "Arial",
                        "color": {"type": "rgb", "rgb": "FFFF0000"},
                    },
                },
                {
                    "text": " italic",
                    "font": {
                        "italic": True,
                        "underline": "single",
                        "color": {"type": "rgb", "rgb": "FF0000FF"},
                    },
                },
            ]
        },
        width=3,
        height=1.5,
        style={"fill_color": "FF4472C4", "outline_color": "FF203864"},
        name="Rich Text Shape",
    )
    session_key, report = _save_validate_close_reload(session_key, output_path)
    assert report["features"]["drawings"] >= 1

    created = next(item for item in _inventory(session_key) if item["name"] == "Rich Text Shape")
    assert created["text"] == "Red italic"
    assert created["rich_text"]["text"] == "Red italic"
    assert created["rich_text"]["runs"] == [
        {
            "text": "Red",
            "start": 0,
            "end": 3,
            "font": {
                "bold": True,
                "size": 14.0,
                "color": {"type": "rgb", "rgb": "FFFF0000"},
                "name": "Arial",
            },
        },
        {
            "text": " italic",
            "start": 3,
            "end": 10,
            "font": {
                "italic": True,
                "underline": "single",
                "color": {"type": "rgb", "rgb": "FF0000FF"},
            },
        },
    ]
    initial_shape = _find_object(_package_state(output_path), "shape", "Rich Text Shape")
    initial_geometry = initial_shape["geometry"]
    initial_dimensions = (initial_shape["cx"], initial_shape["cy"])

    M.excel_update_shape_text(
        session_key,
        "Drawing",
        1,
        rich_text={
            "runs": [
                {
                    "text": "Done",
                    "font": {
                        "strike": True,
                        "color": {"type": "rgb", "rgb": "FF00AA00"},
                    },
                },
                {"text": " plain", "font": {"bold": False}},
            ]
        },
    )
    save_report = json.loads(M.excel_save(
        session_key,
        str(output_path),
        report_format="json",
        verify_preservation=True,
    ))
    assert save_report["verification"]["status"] == "completed"
    assert save_report["verification"]["preservation_ok"] is True
    assert save_report["verification"]["unapproved_difference_count"] == 0
    assert save_report["requested_semantic_paths"][-1] == "sheets/Drawing/drawing_shapes/1/text"
    M.excel_close(session_key)
    session_key = _load_key(M.excel_load(str(output_path)))

    updated = next(item for item in _inventory(session_key) if item["name"] == "Rich Text Shape")
    assert updated["text"] == "Done plain"
    assert updated["rich_text"]["runs"] == [
        {
            "text": "Done",
            "start": 0,
            "end": 4,
            "font": {
                "strike": True,
                "color": {"type": "rgb", "rgb": "FF00AA00"},
            },
        },
        {
            "text": " plain",
            "start": 4,
            "end": 10,
            "font": {"bold": False},
        },
    ]
    updated_shape = _find_object(_package_state(output_path), "shape", "Rich Text Shape")
    assert updated_shape["geometry"] == initial_geometry
    assert (updated_shape["cx"], updated_shape["cy"]) == initial_dimensions
    M.excel_close(session_key)


def test_add_image_rejects_invalid_bytes_before_save(tmp_path):
    session_key, output_path = _create_workbook(tmp_path, "invalid-image.xlsx")
    image_path = tmp_path / "valid-after-invalid.png"
    _write_tiny_png(image_path)
    try:
        with pytest.raises(ValueError, match="unsupported or invalid image bytes"):
            M.excel_add_image(
                session_key,
                "Drawing",
                anchor="A1",
                base64_data="bm90LWEtcmVhbC1pbWFnZQ==",
                mime_type="image/png",
            )

        M.excel_add_image(
            session_key,
            "Drawing",
            anchor="A1",
            source_path=str(image_path),
            name="Valid Image",
        )
        M.excel_save(session_key)
        assert output_path.exists()
    finally:
        M.excel_close(session_key)
