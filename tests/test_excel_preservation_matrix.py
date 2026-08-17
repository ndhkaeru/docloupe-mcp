import hashlib
import json
import os
import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers" / "excel"))

import main as M  # noqa: E402


DEFAULT_FIXTURE_ROOT = Path(r"D:\data-test\excel-preservation-fixtures")
SOURCE_NAMES = (
    "01-audit-87-source.xlsx",
    "02-table-metadata-source.xlsx",
    "03-print-area-source.xlsx",
    "04-rich-text-phonetic-source.xlsx",
    "05-advanced-package-source.xlsm",
    "06-book1-richtext-source.xlsm",
    "07-real-package-source.xlsx",
)


def _fixture_root() -> Path:
    return Path(os.environ.get("DOCLOUPE_EXCEL_FIXTURE_ROOT", DEFAULT_FIXTURE_ROOT))


def _manifest() -> tuple[Path, dict]:
    root = _fixture_root()
    manifest_path = root / "manifest.json"
    if not manifest_path.is_file():
        pytest.skip(f"Excel preservation fixture manifest is unavailable: {manifest_path}")
    return root, json.loads(manifest_path.read_text(encoding="utf-8"))


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _session_key(load_result: str) -> str:
    return load_result.split("session_key='")[1].split("'")[0]


def _package_inventory(session_key: str) -> dict:
    tool = M.mcp._tool_manager._tools["excel_list_package_parts"]
    return json.loads(tool.fn(session_key, max_parts=200, max_relationships_per_part=20))


def _fixture_record(manifest: dict, source_name: str) -> dict:
    for fixture in manifest["fixtures"]:
        if Path(fixture["source"]).name == source_name:
            return fixture
    raise AssertionError(f"Fixture is absent from manifest: {source_name}")


def _assert_manifest_hashes(root: Path, record: dict, source_name: str) -> tuple[Path, Path]:
    source = root / "sources" / source_name
    legacy = Path(record["current_roundtrip"])
    assert source.is_file(), source
    assert legacy.is_file(), legacy
    assert _sha256(source) == record["source_sha256"]
    assert _sha256(legacy) == record["current_roundtrip_sha256"]
    return source, legacy


@pytest.mark.parametrize("source_name", SOURCE_NAMES)
def test_fixture_source_hashes_match_manifest(source_name):
    root, manifest = _manifest()
    record = _fixture_record(manifest, source_name)
    _assert_manifest_hashes(root, record, source_name)


@pytest.mark.parametrize("source_name", SOURCE_NAMES)
def test_legacy_roundtrip_hashes_match_manifest(source_name):
    root, manifest = _manifest()
    record = _fixture_record(manifest, source_name)
    _assert_manifest_hashes(root, record, source_name)


@pytest.mark.parametrize("source_name", SOURCE_NAMES)
def test_no_edit_public_roundtrip_has_no_unapproved_differences(tmp_path, source_name):
    root, manifest = _manifest()
    record = _fixture_record(manifest, source_name)
    source, _ = _assert_manifest_hashes(root, record, source_name)

    source_validation = json.loads(M.excel_validate_workbook(str(source)))
    assert source_validation["valid"] is True
    source_summary = json.loads(M.excel_get_workbook_summary(str(source)))
    output = tmp_path / "roundtrip-output" / source_name
    output.parent.mkdir(parents=True, exist_ok=True)
    session_key = _session_key(M.excel_load(str(source)))
    try:
        source_inventory = _package_inventory(session_key)
        M.excel_save_as_copy(session_key, str(output))
    finally:
        M.excel_close(session_key)

    output_validation = json.loads(M.excel_validate_workbook(str(output)))
    assert output_validation["valid"] is True
    output_session_key = _session_key(M.excel_load(str(output)))
    try:
        output_inventory = _package_inventory(output_session_key)
        output_summary = json.loads(M.excel_get_workbook_summary(str(output)))
    finally:
        M.excel_close(output_session_key)

    package_diff = json.loads(M.excel_diff_package(str(source), str(output)))
    report = json.loads(M.excel_verify_preservation(
        str(output),
        before_path=str(source),
        max_differences=5000,
        fixture_id=source.stem,
    ))
    assert report["unapproved_difference_count"] == 0, json.dumps(report, ensure_ascii=False, indent=2)
    assert report["preservation_ok"] is True
    assert report["files"]["before"]["package_valid"] is True
    assert report["files"]["after"]["package_valid"] is True
    assert {key: value for key, value in source_summary.items() if key != "source"} == {
        key: value for key, value in output_summary.items() if key != "source"
    }
    assert source_inventory["parts"] == output_inventory["parts"]

    evidence = {
        "fixture_id": source.stem,
        "tool_inputs": {
            "excel_validate_workbook": [str(source), str(output)],
            "excel_load": [str(source), str(output)],
            "excel_save_as_copy": {"output_path": str(output)},
            "excel_list_package_parts": {
                "max_parts": 200,
                "max_relationships_per_part": 20,
            },
            "excel_diff_package": {"before_path": str(source), "after_path": str(output)},
            "excel_verify_preservation": {
                "before_path": str(source),
                "after_path": str(output),
                "max_differences": 5000,
            },
        },
        "source": {
            "sha256": _sha256(source),
            "validation": source_validation,
            "summary": source_summary,
            "package_inventory": source_inventory,
        },
        "output": {
            "sha256": _sha256(output),
            "validation": output_validation,
            "summary": output_summary,
            "package_inventory": output_inventory,
        },
        "package_diff": package_diff,
        "verification": report,
    }
    evidence_path = tmp_path / "evidence" / f"{source.stem}.json"
    evidence_path.parent.mkdir(parents=True, exist_ok=True)
    evidence_path.write_text(
        json.dumps(evidence, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    reloaded_evidence = json.loads(evidence_path.read_text(encoding="utf-8"))
    assert reloaded_evidence["verification"]["preservation_ok"] is True
    assert evidence_path.stat().st_size < 2_000_000
