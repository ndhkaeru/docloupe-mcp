from __future__ import annotations

import hashlib
import json
import os
import tempfile
import time
from pathlib import Path
from typing import Any, Callable

from core import reconstruct_excel, serialize_excel
from preservation import (
    _MemorySampler,
    create_excel_backup,
    discard_excel_backup,
    inspect_workbook_pair,
    inspection_pair_performance,
    package_signature_report,
    verify_xlsx_preservation,
)


class SaveTransactionError(RuntimeError):
    def __init__(self, details: dict[str, Any]):
        self.details = details
        super().__init__(json.dumps(details, ensure_ascii=False, sort_keys=True, default=str))


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_state(path: str | Path) -> dict[str, Any]:
    resolved = Path(path).expanduser().resolve()
    if not resolved.is_file():
        return {"exists": False, "size": None, "sha256": None}
    return {
        "exists": True,
        "size": resolved.stat().st_size,
        "sha256": sha256_file(resolved),
    }


def create_staging_path(destination: str | Path) -> Path:
    resolved = Path(destination).expanduser().resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(
        prefix=f".{resolved.stem}.docloupe-",
        suffix=resolved.suffix,
        dir=resolved.parent,
    )
    os.close(handle)
    staging = Path(temporary)
    staging.unlink(missing_ok=True)
    return staging


def remove_staging_path(path: str | Path | None) -> bool:
    if path is None:
        return True
    staging = Path(path)
    try:
        staging.unlink(missing_ok=True)
    except OSError:
        return False
    saving_temporary = Path(str(staging) + ".~saving.tmp")
    try:
        saving_temporary.unlink(missing_ok=True)
    except OSError:
        return False
    return not staging.exists() and not saving_temporary.exists()


def verification_reference(session_data: dict, destination: str | Path) -> str | None:
    destination_path = Path(destination).expanduser().resolve()
    baseline = session_data.get("_verification_baseline_path")
    if baseline:
        baseline_path = Path(baseline).expanduser().resolve()
    else:
        source = Path(session_data["source"]).expanduser().resolve()
        default_output = Path(
            session_data.get("_default_output_path") or source
        ).expanduser().resolve()
        if session_data.get("_new_workbook") and source == default_output:
            baseline_path = None
        else:
            baseline_path = source
    if baseline_path is None:
        return None
    if baseline_path == destination_path:
        return str(destination_path) if destination_path.is_file() else None
    return str(baseline_path) if baseline_path.is_file() else None


def _filtered_merge_reference(session_data: dict) -> Path:
    baseline = session_data.get("_verification_baseline_path")
    if baseline:
        baseline_path = Path(baseline).expanduser().resolve()
        if baseline_path.is_file():
            return baseline_path
    source = Path(session_data["source"]).expanduser().resolve()
    if source.is_file():
        return source
    raise ValueError(
        "Session was loaded with a sheet_name filter and no readable baseline "
        "exists to merge unloaded sheets back."
    )


def _merge_filtered_session(session_data: dict) -> tuple[dict, bool]:
    if not session_data.get("_sheet_filter"):
        return session_data, False
    reference = _filtered_merge_reference(session_data)
    try:
        full = serialize_excel(str(reference))
    except Exception as exc:
        raise ValueError(
            "Session was loaded with a sheet_name filter and the baseline file "
            f"can no longer be read to merge unloaded sheets back ({exc})."
        ) from exc
    loaded = set(session_data.get("_loaded_disk_names") or [])
    merged: list[dict] = []
    spliced = False
    for sheet in full["sheets"]:
        if sheet["name"] in loaded:
            if not spliced:
                merged.extend(session_data["sheets"])
                spliced = True
        else:
            merged.append(sheet)
    if not spliced:
        merged.extend(session_data["sheets"])
    names = [sheet["name"] for sheet in merged]
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise ValueError(
            f"Sheet name collision while merging the filtered session back: {duplicates}. "
            "Rename the session sheet or save to a different output_path."
        )
    combined = {**full, **session_data, "source": str(reference), "sheets": merged}
    return combined, True


def _verification_summary(
    report: dict,
    reference_path: str,
    requested_paths: list[str],
) -> dict:
    changes = report.get("changes") or []
    return {
        "status": "completed",
        "reference_path": reference_path,
        "preservation_ok": report.get("preservation_ok"),
        "equivalent": report.get("equivalent"),
        "change_count": report.get("change_count", 0),
        "unapproved_difference_count": report.get("unapproved_difference_count", 0),
        "classification_counts": report.get("classification_counts") or {},
        "severity_counts": report.get("severity_counts") or {},
        "requested_paths": requested_paths,
        "changed_semantic_paths": [
            item.get("path") for item in changes if item.get("path")
        ],
        "truncated": bool(report.get("truncated")),
        "recommendation": report.get("recommendation"),
    }


def execute_save_stage(
    session_data: dict,
    *,
    staging_path: str,
    verification_reference_path: str | None,
    verify_preservation: bool,
    max_differences: int,
    requested_paths: list[str],
    intentional_edit: bool,
    reconstruct: Callable[[dict, str], list[str] | None] = reconstruct_excel,
    verify: Callable[..., dict] = verify_xlsx_preservation,
    signature_report: Callable[..., dict] = package_signature_report,
) -> dict[str, Any]:
    staging = Path(staging_path).expanduser().resolve()
    to_write, sheet_filter_merged = _merge_filtered_session(session_data)
    warnings = list(reconstruct(to_write, str(staging)) or [])

    left_inspection = None
    right_inspection = None
    metadata_seconds = 0.0
    semantic_verification_seconds = 0.0
    memory_sampler = None
    memory_metrics: dict[str, int | None] = {}
    inspection_started_at = 0.0
    use_shared_inspection = bool(
        verify_preservation
        and verification_reference_path
        and (
            verify is verify_xlsx_preservation
            or signature_report is package_signature_report
        )
    )
    if use_shared_inspection:
        memory_sampler = _MemorySampler()
        memory_sampler.start()
        inspection_started_at = time.perf_counter()
        metadata_started_at = time.perf_counter()
        try:
            left_inspection, right_inspection = inspect_workbook_pair(
                verification_reference_path,
                staging,
            )
        except Exception:
            memory_sampler.finish()
            raise
        metadata_seconds = time.perf_counter() - metadata_started_at

    try:
        signature_kwargs = {}
        if signature_report is package_signature_report and left_inspection is not None:
            signature_kwargs = {
                "before_inspection": left_inspection,
                "after_inspection": right_inspection,
            }
        try:
            package_signatures = signature_report(
                verification_reference_path,
                staging,
                intentional_edit=bool(intentional_edit),
                **signature_kwargs,
            )
        except Exception as exc:
            package_signatures = {
                "present": None,
                "intentional_edit": bool(intentional_edit),
                "status": "inspection_error",
                "parts_before": {},
                "parts_after": {},
                "parts_preserved": False,
                "message": (
                    "Could not inspect OOXML package signature parts after save: "
                    f"{exc}"
                ),
            }
        if package_signatures["status"] in {
            "requires_resigning",
            "signature_parts_changed",
            "inspection_error",
        }:
            warnings.append(package_signatures["message"])

        if verify_preservation and verification_reference_path:
            verify_kwargs = {}
            if verify is verify_xlsx_preservation and left_inspection is not None:
                verify_kwargs = {
                    "before_inspection": left_inspection,
                    "after_inspection": right_inspection,
                }
            semantic_started_at = time.perf_counter()
            verifier_report = verify(
                verification_reference_path,
                str(staging),
                int(max_differences),
                requested_paths=requested_paths,
                **verify_kwargs,
            )
            semantic_verification_seconds = time.perf_counter() - semantic_started_at
            verification = _verification_summary(
                verifier_report,
                verification_reference_path,
                requested_paths,
            )
        elif verify_preservation:
            verification = {
                "status": "skipped",
                "reference_path": None,
                "requested_paths": requested_paths,
                "reason": "No pre-save semantic reference exists for this new workbook yet.",
            }
        else:
            verification = {
                "status": "not_run",
                "reference_path": verification_reference_path,
                "requested_paths": requested_paths,
            }
    finally:
        if left_inspection is not None:
            left_inspection.release_raw_parts()
        if right_inspection is not None:
            right_inspection.release_raw_parts()
        if memory_sampler is not None:
            memory_metrics = memory_sampler.finish()

    inspection_performance = None
    if left_inspection is not None and right_inspection is not None:
        inspection_performance = inspection_pair_performance(
            left_inspection,
            right_inspection,
            metadata_seconds=metadata_seconds,
            semantic_verification_seconds=semantic_verification_seconds,
            total_tool_seconds=time.perf_counter() - inspection_started_at,
            memory=memory_metrics,
        )
        verification["performance"] = inspection_performance

    return {
        "file_size": staging.stat().st_size,
        "sheet_count": len(to_write["sheets"]),
        "total_rows": sum(len(sheet["rows"]) for sheet in to_write["sheets"]),
        "warnings": warnings,
        "package_signatures": package_signatures,
        "verification": verification,
        "inspection_performance": inspection_performance,
        "sheet_filter_merged": sheet_filter_merged,
    }


def require_preservation_success(result: dict[str, Any]) -> None:
    verification = result.get("verification") or {}
    if (
        verification.get("status") == "completed"
        and verification.get("preservation_ok") is False
    ):
        raise SaveTransactionError({
            "code": "EXCEL_SAVE_PRESERVATION_FAILED",
            "message": "Preservation verification rejected the staged workbook.",
            "verification": verification,
        })


def _assert_state(path: Path, expected: dict[str, Any], label: str) -> None:
    actual = file_state(path)
    if actual != expected:
        raise SaveTransactionError({
            "code": "EXCEL_SAVE_CONCURRENT_FILE_CHANGE",
            "message": f"{label} changed while the save transaction was running.",
            "path": str(path),
            "expected": expected,
            "actual": actual,
        })


def commit_staging_file(
    *,
    staging_path: str | Path,
    destination: str | Path,
    expected_destination_state: dict[str, Any],
    verification_reference_path: str | None,
    expected_reference_state: dict[str, Any] | None,
) -> dict | None:
    staging = Path(staging_path).expanduser().resolve()
    destination_path = Path(destination).expanduser().resolve()
    if not staging.is_file():
        raise SaveTransactionError({
            "code": "EXCEL_SAVE_STAGING_MISSING",
            "message": "The staged workbook is missing before commit.",
            "staging_path": str(staging),
        })
    _assert_state(destination_path, expected_destination_state, "Destination")
    if verification_reference_path and expected_reference_state is not None:
        reference = Path(verification_reference_path).expanduser().resolve()
        if reference != destination_path:
            _assert_state(reference, expected_reference_state, "Verification reference")

    backup = None
    if expected_destination_state.get("exists"):
        backup = create_excel_backup(str(destination_path), str(destination_path))
        try:
            if backup.get("sha256") != expected_destination_state.get("sha256"):
                raise SaveTransactionError({
                    "code": "EXCEL_SAVE_BACKUP_HASH_MISMATCH",
                    "message": "The pre-save backup does not match the destination.",
                    "destination": str(destination_path),
                    "destination_sha256": expected_destination_state.get("sha256"),
                    "backup_sha256": backup.get("sha256"),
                })
            _assert_state(destination_path, expected_destination_state, "Destination")
        except Exception:
            discard_excel_backup(backup)
            raise

    try:
        os.replace(staging, destination_path)
    except Exception:
        discard_excel_backup(backup)
        raise
    return backup
