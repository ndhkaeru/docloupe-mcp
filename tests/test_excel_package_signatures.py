import hashlib
import json
import os
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl

sys.path.insert(0, str(Path(__file__).parents[1] / "servers" / "excel"))
import main as M


_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_SIGNATURE_XML = (
    b'<Signature xmlns="http://www.w3.org/2000/09/xmldsig#">'
    b"<SignedInfo/><SignatureValue>QVVESVQ=</SignatureValue></Signature>"
)


def _session_key(result: str) -> str:
    return result.split("session_key='")[1].split("'")[0]


def _write_signed_workbook(path: Path) -> None:
    workbook = openpyxl.Workbook()
    workbook.active.title = "Sheet"
    workbook.active["A1"] = "original"
    workbook.active["B1"] = "keep"
    workbook.save(path)
    workbook.close()

    rebuilt = path.with_suffix(path.suffix + ".signed.tmp")
    with ZipFile(path, "r") as source, ZipFile(rebuilt, "w") as target:
        for info in source.infolist():
            raw = source.read(info.filename)
            if info.filename == "[Content_Types].xml":
                root = ET.fromstring(raw)
                ET.SubElement(root, f"{{{_CONTENT_TYPES_NS}}}Override", {
                    "PartName": "/_xmlsignatures/origin.sigs",
                    "ContentType": "application/vnd.openxmlformats-package.digital-signature-origin",
                })
                ET.SubElement(root, f"{{{_CONTENT_TYPES_NS}}}Override", {
                    "PartName": "/_xmlsignatures/sig1.xml",
                    "ContentType": "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                })
                raw = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            target.writestr(info, raw)
        target.writestr("_xmlsignatures/origin.sigs", b"", compress_type=ZIP_DEFLATED)
        target.writestr("_xmlsignatures/sig1.xml", _SIGNATURE_XML, compress_type=ZIP_DEFLATED)
    os.replace(rebuilt, path)


def _signature_hashes(path: Path) -> dict[str, str]:
    with ZipFile(path, "r") as archive:
        return {
            name: hashlib.sha256(archive.read(name)).hexdigest()
            for name in sorted(archive.namelist())
            if name.startswith("_xmlsignatures/") and not name.endswith("/")
        }


def test_no_edit_save_as_copy_preserves_signature_parts_without_invalidation_warning(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCLOUPE_EXCEL_BACKUP_DIR", str(tmp_path / "backups"))
    source = tmp_path / "signed-source.xlsx"
    target = tmp_path / "signed-copy.xlsx"
    _write_signed_workbook(source)
    before_hashes = _signature_hashes(source)

    session_key = _session_key(M.excel_load(str(source)))
    report = json.loads(M.excel_save_as_copy(session_key, str(target), report_format="json"))

    signatures = report["package_signatures"]
    assert signatures["status"] == "preserved_unverified"
    assert signatures["present"] is True
    assert signatures["intentional_edit"] is False
    assert signatures["parts_preserved"] is True
    assert signatures["parts_before"] == before_hashes
    assert signatures["parts_after"] == before_hashes
    assert signatures["message"] not in report["warnings"]
    assert _signature_hashes(target) == before_hashes
    assert json.loads(M.excel_validate_workbook(str(target)))["valid"] is True


def test_intentional_cell_edit_preserves_parts_and_requires_resigning(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCLOUPE_EXCEL_BACKUP_DIR", str(tmp_path / "backups"))
    source = tmp_path / "signed-source.xlsx"
    target = tmp_path / "signed-edited.xlsx"
    _write_signed_workbook(source)
    before_hashes = _signature_hashes(source)

    session_key = _session_key(M.excel_load(str(source)))
    M.excel_edit_cells(session_key, "Sheet", [{"row_index": 0, "edits": {0: "edited"}}])
    report = json.loads(M.excel_save_as_copy(session_key, str(target), report_format="json"))

    signatures = report["package_signatures"]
    assert signatures["status"] == "requires_resigning"
    assert signatures["intentional_edit"] is True
    assert signatures["parts_preserved"] is True
    assert signatures["parts_before"] == before_hashes
    assert signatures["parts_after"] == before_hashes
    assert signatures["message"] in report["warnings"]
    assert "signed again" in signatures["message"]
    assert _signature_hashes(target) == before_hashes
    workbook = openpyxl.load_workbook(target)
    assert workbook.active["A1"].value == "edited"
    assert workbook.active["B1"].value == "keep"
    workbook.close()


def test_signature_part_mutation_is_reported_as_critical_change(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCLOUPE_EXCEL_BACKUP_DIR", str(tmp_path / "backups"))
    source = tmp_path / "signed-source.xlsx"
    target = tmp_path / "signed-mutated.xlsx"
    _write_signed_workbook(source)

    session_key = _session_key(M.excel_load(str(source)))
    M._package_tools["excel_upsert_package_part"](
        session_key,
        "_xmlsignatures/sig1.xml",
        '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#">'
        "<SignedInfo/><SignatureValue>Q0hBTkdFRA==</SignatureValue></Signature>",
        content_type="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
    )
    report = json.loads(M.excel_save_as_copy(session_key, str(target), report_format="json"))

    signatures = report["package_signatures"]
    assert signatures["status"] == "signature_parts_changed"
    assert signatures["intentional_edit"] is True
    assert signatures["parts_preserved"] is False
    assert signatures["parts_before"] != signatures["parts_after"]
    assert signatures["message"].startswith("CRITICAL:")
    assert signatures["message"] in report["warnings"]
    with ZipFile(target, "r") as archive:
        content_types = ET.fromstring(archive.read("[Content_Types].xml"))
    signature_overrides = [
        node for node in content_types
        if node.attrib.get("PartName") == "/_xmlsignatures/sig1.xml"
    ]
    assert len(signature_overrides) == 1
    assert json.loads(M.excel_validate_workbook(str(target)))["valid"] is True
