import sys
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers" / "excel"))

import core  # noqa: E402
import main as M  # noqa: E402


_BASELINE_KEYS = (
    "_baseline_content_hash",
    "_baseline_style_hash",
    "_baseline_structure_hash",
)


def _write_workbook(path: Path, *, columns: int = 2) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet"
    for column in range(1, columns + 1):
        worksheet.cell(row=1, column=column, value=f"value-{column}")
    workbook.save(path)
    workbook.close()


def _session_key(result: str) -> str:
    return result.split("session_key=", 1)[1].split(" |", 1)[0].strip("'")


def test_serialize_excel_does_not_create_eager_cell_baselines(tmp_path, monkeypatch):
    source = tmp_path / "no-eager-baseline.xlsx"
    _write_workbook(source, columns=128)

    def unexpected_baseline(_cell):
        raise AssertionError("serialize_excel must not baseline every cell")

    monkeypatch.setattr(core, "_cell_baseline", unexpected_baseline)
    data = core.serialize_excel(str(source))

    cells = data["sheets"][0]["rows"][0]["cells"]
    assert len(cells) == 128
    assert all(not any(key in cell for key in _BASELINE_KEYS) for cell in cells)


def test_initial_load_semantic_digest_count_does_not_scale_per_cell(tmp_path, monkeypatch):
    source = tmp_path / "digest-count.xlsx"
    _write_workbook(source, columns=128)

    original_digest = core._semantic_digest
    digest_calls = 0

    def counted_digest(value):
        nonlocal digest_calls
        digest_calls += 1
        return original_digest(value)

    monkeypatch.setattr(core, "_semantic_digest", counted_digest)
    core.serialize_excel(str(source))

    assert digest_calls < 20


def test_cell_baseline_is_captured_only_once():
    cell = {
        "v": "original",
        "data_type": "inlineStr",
        "present": True,
        "merge": {},
        "fill": None,
        "bold": False,
    }

    core._cell_baseline(cell)
    original_hashes = tuple(cell[key] for key in _BASELINE_KEYS)
    cell["v"] = "second"
    core._cell_baseline(cell)

    assert tuple(cell[key] for key in _BASELINE_KEYS) == original_hashes
    assert core._current_cell_group_hashes(cell)[0] != original_hashes[0]


def test_public_edit_multiple_times_then_revert_uses_verified_exact_copy(tmp_path):
    source = tmp_path / "edit-revert.xlsx"
    output = tmp_path / "edit-revert-copy.xlsx"
    _write_workbook(source, columns=2)
    original_bytes = source.read_bytes()

    session_key = _session_key(M.excel_load(str(source)))
    try:
        data = M._sessions[M._resolve_session_key(session_key)]
        cell = data["sheets"][0]["rows"][0]["cells"][0]
        original_type = cell["data_type"]
        assert not any(key in cell for key in _BASELINE_KEYS)

        M.excel_edit_cells(
            session_key,
            "Sheet",
            [{"row_index": 0, "edits": {0: {"value": "first", "data_type": original_type}}}],
        )
        captured_hashes = tuple(cell[key] for key in _BASELINE_KEYS)

        M.excel_edit_cells(
            session_key,
            "Sheet",
            [{"row_index": 0, "edits": {0: {"value": "second", "data_type": original_type}}}],
        )
        assert tuple(cell[key] for key in _BASELINE_KEYS) == captured_hashes

        M.excel_edit_cells(
            session_key,
            "Sheet",
            [{"row_index": 0, "edits": {0: {"value": "value-1", "data_type": original_type}}}],
        )
        assert core._content_only_changes(data) == {}

        core.reconstruct_excel(data, str(output))
        assert output.read_bytes() == original_bytes
    finally:
        M._sessions.pop(M._resolve_session_key(session_key), None)


def test_fill_and_clear_capture_baselines_and_dirty_paths(tmp_path):
    source = tmp_path / "fill-clear.xlsx"
    _write_workbook(source, columns=2)

    session_key = _session_key(M.excel_load(str(source)))
    try:
        data = M._sessions[M._resolve_session_key(session_key)]
        first = data["sheets"][0]["rows"][0]["cells"][0]
        second = data["sheets"][0]["rows"][0]["cells"][1]

        M.excel_fill_column(session_key, "Sheet", 0, 0, 0, value="filled")
        M.excel_clear_range(session_key, "Sheet", 0, 1, 0, 1)

        assert all(key in first for key in _BASELINE_KEYS)
        assert all(key in second for key in _BASELINE_KEYS)
        assert "cells" in data["_dirty_features"]
        assert "sheets/Sheet/cells/A1" in data["_dirty_paths"]
        assert "sheets/Sheet/cells/B1" in data["_dirty_paths"]
    finally:
        M._sessions.pop(M._resolve_session_key(session_key), None)
