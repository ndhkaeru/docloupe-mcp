"""Targeted tests for the expert OOXML package edit API."""

from __future__ import annotations

import base64
import copy
import hashlib
import json
import sys
import zipfile
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers" / "excel"))

import package_tools as P  # noqa: E402


class FakeMCP:
    def __init__(self) -> None:
        self.tools = {}

    def tool(self):
        def decorator(function):
            self.tools[function.__name__] = function
            return function

        return decorator


def _write_minimal_xlsx(path: Path) -> None:
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''
    root_relationships = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
    workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''
    workbook_relationships = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''
    worksheet = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'''
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_relationships)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_relationships)
        archive.writestr("xl/worksheets/sheet1.xml", worksheet)


def _make_session(tmp_path: Path):
    source = tmp_path / "source.xlsx"
    _write_minimal_xlsx(source)
    sessions = {"session": {"source": str(source), "sheets": []}}
    mcp = FakeMCP()
    registered = P.register_package_tools(mcp, sessions.__getitem__)
    return source, sessions["session"], registered


def _call(function, *args, **kwargs):
    return json.loads(function(*args, **kwargs))


def _relationship(
    relationship_id: str,
    target: str,
    target_mode: str = "Internal",
    relationship_type: str = "urn:docloupe:test:relationship",
):
    return {
        "id": relationship_id,
        "type": relationship_type,
        "target": target,
        "target_mode": target_mode,
    }


def test_registers_complete_public_package_api(tmp_path):
    _, _, registered = _make_session(tmp_path)

    assert set(registered) == {
        "excel_list_package_parts",
        "excel_read_package_part",
        "excel_upsert_package_part",
        "excel_delete_package_part",
        "excel_set_package_relationships",
        "excel_set_package_content_types",
        "excel_apply_package_transaction",
    }


def test_transaction_stages_xml_binary_relationships_and_content_types(tmp_path):
    source, session, tools = _make_session(tmp_path)
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    xml_content = '<custom xmlns="urn:docloupe:test"><value>created by tool</value></custom>'
    binary_content = b"\x00\x01\x02docloupe\xff"

    report = _call(
        tools["excel_apply_package_transaction"],
        "session",
        upsert=[
            {
                "path": "customXml/item1.xml",
                "content": xml_content,
                "encoding": "text",
                "content_type": "application/vnd.docloupe.custom+xml",
            },
            {
                "path": "xl/embeddings/blob.bin",
                "content": base64.b64encode(binary_content).decode("ascii"),
                "encoding": "base64",
                "content_type": "application/vnd.docloupe.binary",
            },
        ],
        relationships=[
            {
                "source_part": "customXml/item1.xml",
                "relationships": [
                    _relationship("rIdInternal", "../xl/embeddings/blob.bin"),
                    _relationship(
                        "rIdExternal",
                        "https://example.test/reference",
                        target_mode="External",
                    ),
                ],
            }
        ],
    )

    assert report["mutated"] is True
    assert report["dirty"] == {
        "upsert_count": 2,
        "upsert_bytes": len(xml_content.encode("utf-8")) + len(binary_content),
        "delete_count": 0,
        "relationship_source_count": 1,
        "relationship_count": 2,
        "content_type_default_edits": 0,
        "content_type_override_edits": 2,
        "dirty": True,
    }
    assert hashlib.sha256(source.read_bytes()).hexdigest() == source_hash

    edits = session["_package_edits"]
    assert set(edits) == {"upsert", "delete", "relationships", "content_types"}
    assert edits["delete"] == []
    assert edits["relationships"]["customXml/item1.xml"] == [
        {
            "id": "rIdInternal",
            "type": "urn:docloupe:test:relationship",
            "target": "../xl/embeddings/blob.bin",
            "target_mode": "Internal",
        },
        {
            "id": "rIdExternal",
            "type": "urn:docloupe:test:relationship",
            "target": "https://example.test/reference",
            "target_mode": "External",
        },
    ]
    assert edits["content_types"] == {
        "defaults": {},
        "overrides": {
            "customXml/item1.xml": "application/vnd.docloupe.custom+xml",
            "xl/embeddings/blob.bin": "application/vnd.docloupe.binary",
        },
    }
    assert base64.b64decode(edits["upsert"]["xl/embeddings/blob.bin"]["data_base64"]) == binary_content

    listing = _call(tools["excel_list_package_parts"], "session", max_parts=100)
    parts = {item["path"]: item for item in listing["parts"]}
    assert parts["customXml/item1.xml"]["content_type"] == "application/vnd.docloupe.custom+xml"
    assert parts["customXml/item1.xml"]["outbound_relationship_count"] == 2
    assert parts["xl/embeddings/blob.bin"]["content_type"] == "application/vnd.docloupe.binary"
    assert parts["xl/embeddings/blob.bin"]["inbound_relationship_count"] == 1
    assert parts["customXml/_rels/item1.xml.rels"]["content_type"] == (
        "application/vnd.openxmlformats-package.relationships+xml"
    )

    xml_read = _call(
        tools["excel_read_package_part"],
        "session",
        "customXml/item1.xml",
        output_mode="xml",
    )
    assert xml_read["content"] == xml_content
    assert xml_read["truncated"] is False

    binary_read = _call(
        tools["excel_read_package_part"],
        "session",
        "xl/embeddings/blob.bin",
        output_mode="base64",
    )
    assert base64.b64decode(binary_read["content"]) == binary_content
    raw_read = _call(
        tools["excel_read_package_part"],
        "session",
        "xl/embeddings/blob.bin",
        output_mode="raw",
    )
    assert raw_read["content"] == list(binary_content)


def test_package_edit_verifier_patterns_are_granular_and_effective(tmp_path):
    source, session, tools = _make_session(tmp_path)
    _call(
        tools["excel_apply_package_transaction"],
        "session",
        upsert=[
            {
                "path": "customXml/item1.xml",
                "content": "<custom/>",
                "content_type": "application/vnd.docloupe.custom+xml",
            }
        ],
        relationships=[
            {
                "source_part": "customXml/item1.xml",
                "relationships": [_relationship("rId1", "../xl/workbook.xml")],
            }
        ],
        content_types={
            "defaults": [
                {
                    "extension": "foo",
                    "content_type": "application/vnd.docloupe.foo",
                }
            ]
        },
    )

    assert P.package_edit_verifier_patterns(session, source) == [
        "package/customXml/item1.xml",
        "package/customXml/_rels/item1.xml.rels",
        "package/relationships/customXml/_rels/item1.xml.rels#*",
        "package/content_types/Default:foo",
        "package/content_types/Override:/customXml/item1.xml",
    ]


def test_package_edit_verifier_patterns_drop_cancelled_edits(tmp_path):
    source, session, tools = _make_session(tmp_path)
    _call(
        tools["excel_upsert_package_part"],
        "session",
        "customXml/transient.xml",
        "<transient/>",
        content_type="application/vnd.docloupe.transient+xml",
    )
    _call(
        tools["excel_delete_package_part"],
        "session",
        "customXml/transient.xml",
    )

    assert P.package_edit_verifier_patterns(session, source) == []


def test_relationship_and_content_type_setters_store_apply_ready_records(tmp_path):
    _, session, tools = _make_session(tmp_path)
    _call(
        tools["excel_apply_package_transaction"],
        "session",
        upsert=[
            {
                "path": "customXml/source.xml",
                "content": "<source/>",
                "content_type": "application/vnd.docloupe.source+xml",
            },
            {
                "path": "customXml/target.dat",
                "content": base64.b64encode(b"target").decode("ascii"),
                "encoding": "base64",
                "content_type": "application/vnd.docloupe.target",
            },
        ],
    )

    relationship_report = _call(
        tools["excel_set_package_relationships"],
        "session",
        "customXml/source.xml",
        [
            _relationship("rel1", "target.dat"),
            _relationship("rel2", "https://example.test/external", "External"),
        ],
    )
    assert relationship_report["dirty"]["relationship_count"] == 2

    content_type_report = _call(
        tools["excel_set_package_content_types"],
        "session",
        defaults=[
            {
                "extension": ".dat",
                "content_type": "application/vnd.docloupe.default-data",
            }
        ],
        overrides=[
            {
                "part_path": "customXml/target.dat",
                "content_type": None,
            }
        ],
    )
    assert content_type_report["dirty"]["content_type_default_edits"] == 1
    assert session["_package_edits"]["relationships"]["customXml/source.xml"][0] == {
        "id": "rel1",
        "type": "urn:docloupe:test:relationship",
        "target": "target.dat",
        "target_mode": "Internal",
    }
    assert session["_package_edits"]["content_types"] == {
        "defaults": {"dat": "application/vnd.docloupe.default-data"},
        "overrides": {
            "customXml/source.xml": "application/vnd.docloupe.source+xml",
            "customXml/target.dat": None,
        },
    }

    listing = _call(tools["excel_list_package_parts"], "session", prefix="customXml/")
    parts = {item["path"]: item for item in listing["parts"]}
    assert parts["customXml/target.dat"]["content_type"] == "application/vnd.docloupe.default-data"
    assert parts["customXml/target.dat"]["inbound_relationships"][0]["target_mode"] == "Internal"


def test_read_and_list_outputs_are_bounded(tmp_path):
    _, _, tools = _make_session(tmp_path)
    content = "<root><value>" + ("x" * 1000) + "</value></root>"
    _call(
        tools["excel_upsert_package_part"],
        "session",
        "customXml/large.xml",
        content,
        content_type="application/vnd.docloupe.large+xml",
    )

    first = _call(
        tools["excel_read_package_part"],
        "session",
        "customXml/large.xml",
        output_mode="xml",
        max_bytes=32,
    )
    assert first["returned_bytes"] == 32
    assert first["truncated"] is True
    assert len(first["content"].encode("utf-8")) == 32

    second = _call(
        tools["excel_read_package_part"],
        "session",
        "customXml/large.xml",
        output_mode="text",
        offset=32,
        max_bytes=17,
    )
    assert second["offset"] == 32
    assert second["returned_bytes"] == 17

    listing = _call(tools["excel_list_package_parts"], "session", max_parts=1)
    assert listing["returned_part_count"] == 1
    assert listing["truncated"] is True


def test_invalid_inputs_are_atomic_and_leave_session_unchanged(tmp_path):
    _, session, tools = _make_session(tmp_path)
    apply_transaction = tools["excel_apply_package_transaction"]

    invalid_transactions = [
        {
            "upsert": [
                {
                    "path": "../escape.xml",
                    "content": "<root/>",
                    "content_type": "application/xml",
                }
            ]
        },
        {
            "upsert": [
                {
                    "path": "customXml/broken.xml",
                    "content": "<root>",
                    "content_type": "application/xml",
                }
            ]
        },
        {
            "upsert": [
                {
                    "path": "xl/embeddings/broken.bin",
                    "content": "not base64!",
                    "encoding": "base64",
                    "content_type": "application/octet-stream",
                }
            ]
        },
        {
            "relationships": [
                {
                    "source_part": "xl/workbook.xml",
                    "relationships": [
                        _relationship("duplicate", "worksheets/sheet1.xml"),
                        _relationship("duplicate", "worksheets/sheet1.xml"),
                    ],
                }
            ]
        },
        {
            "relationships": [
                {
                    "source_part": "xl/workbook.xml",
                    "relationships": [_relationship("missing", "missing-part.xml")],
                }
            ]
        },
        {
            "content_types": {
                "defaults": [
                    {"extension": "dup", "content_type": "application/one"},
                    {"extension": ".DUP", "content_type": "application/two"},
                ]
            }
        },
        {
            "content_types": {
                "overrides": [
                    {"part_path": "customXml/same.xml", "content_type": "application/one+xml"},
                    {"part_path": "/customXml/same.xml", "content_type": "application/two+xml"},
                ]
            }
        },
        {
            "upsert": [
                {
                    "path": "[Content_Types].xml",
                    "content": "<Types/>",
                    "content_type": "application/xml",
                }
            ]
        },
    ]

    for transaction in invalid_transactions:
        before = copy.deepcopy(session)
        with pytest.raises(P.PackageToolError):
            apply_transaction("session", **transaction)
        assert session == before


def test_late_dangling_validation_rolls_back_other_valid_operations(tmp_path):
    _, session, tools = _make_session(tmp_path)
    before = copy.deepcopy(session)

    with pytest.raises(P.PackageToolError, match="Dangling internal relationship"):
        tools["excel_apply_package_transaction"](
            "session",
            upsert=[
                {
                    "path": "customXml/valid.xml",
                    "content": "<valid/>",
                    "content_type": "application/vnd.docloupe.valid+xml",
                }
            ],
            relationships=[
                {
                    "source_part": "customXml/valid.xml",
                    "relationships": [_relationship("rId1", "missing.bin")],
                }
            ],
        )

    assert session == before


def test_delete_rejects_inbound_edge_then_succeeds_after_graph_update(tmp_path):
    source, session, tools = _make_session(tmp_path)
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    _call(
        tools["excel_apply_package_transaction"],
        "session",
        upsert=[
            {
                "path": "customXml/source.xml",
                "content": "<source/>",
                "content_type": "application/vnd.docloupe.source+xml",
            },
            {
                "path": "customXml/target.bin",
                "content": base64.b64encode(b"target").decode("ascii"),
                "encoding": "base64",
                "content_type": "application/vnd.docloupe.target",
            },
        ],
        relationships=[
            {
                "source_part": "customXml/source.xml",
                "relationships": [_relationship("rId1", "target.bin")],
            }
        ],
    )

    before_failed_delete = copy.deepcopy(session)
    with pytest.raises(P.PackageToolError, match="Dangling internal relationship"):
        tools["excel_delete_package_part"]("session", "customXml/target.bin")
    assert session == before_failed_delete

    _call(
        tools["excel_set_package_relationships"],
        "session",
        "customXml/source.xml",
        [_relationship("external", "https://example.test/target", "External")],
    )
    delete_report = _call(
        tools["excel_delete_package_part"],
        "session",
        "customXml/target.bin",
    )

    assert delete_report["mutated"] is True
    assert "customXml/target.bin" not in session["_package_edits"]["upsert"]
    assert session["_package_edits"]["content_types"]["overrides"]["customXml/target.bin"] is None
    with pytest.raises(P.PackageToolError, match="Package part not found"):
        tools["excel_read_package_part"]("session", "customXml/target.bin")
    assert hashlib.sha256(source.read_bytes()).hexdigest() == source_hash


def test_transaction_can_delete_existing_part_with_relationship_and_content_type_cleanup(tmp_path):
    _, session, tools = _make_session(tmp_path)
    _call(
        tools["excel_upsert_package_part"],
        "session",
        "customXml/delete-me.xml",
        "<delete-me/>",
        content_type="application/vnd.docloupe.delete+xml",
    )

    delete_report = _call(
        tools["excel_delete_package_part"],
        "session",
        "customXml/delete-me.xml",
    )

    assert delete_report["change_count"] >= 1
    assert "customXml/delete-me.xml" not in session["_package_edits"]["upsert"]
    assert session["_package_edits"]["content_types"]["overrides"]["customXml/delete-me.xml"] is None
    listing = _call(
        tools["excel_list_package_parts"],
        "session",
        prefix="customXml/delete-me.xml",
        include_deleted=True,
    )
    assert listing["part_count"] == 0


def test_main_server_registers_public_package_api():
    import main as excel_main

    registered = set(excel_main.mcp._tool_manager._tools)
    assert {
        "excel_list_package_parts",
        "excel_read_package_part",
        "excel_upsert_package_part",
        "excel_delete_package_part",
        "excel_set_package_relationships",
        "excel_set_package_content_types",
        "excel_apply_package_transaction",
    } <= registered
