from __future__ import annotations

import hashlib
import json
import math
from collections import Counter
import os
import posixpath
import re
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Callable

from core import reconstruct_excel, serialize_excel
from preservation import (
    _MemorySampler,
    _workbook_snapshot,
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


def _guard_advanced_package_parts(reference_path: str | None, staging: Path, allowed_deletions: set[str]) -> None:
    if not reference_path or not zipfile.is_zipfile(reference_path):
        return
    protected_prefixes = (
        "xl/drawings/", "xl/charts/", "xl/media/", "xl/printerSettings/",
        "xl/vbaProject", "xl/externalLinks/", "xl/pivot", "xl/slicer",
        "xl/activeX/", "xl/embeddings/", "_xmlsignatures/",
    )
    protected_parts = {"xl/calcChain.xml"}
    with zipfile.ZipFile(reference_path) as before, zipfile.ZipFile(staging) as after:
        old_parts = set(before.namelist())
        new_parts = set(after.namelist())
        missing = sorted(
            part for part in old_parts - new_parts
            if part not in allowed_deletions and (part in protected_parts or part.startswith(protected_prefixes))
        )
        macro_parts = {part for part in old_parts if part.startswith("xl/vbaProject")}
        if macro_parts:
            namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
            old_types = ET.fromstring(before.read("[Content_Types].xml"))
            new_types = ET.fromstring(after.read("[Content_Types].xml"))
            def workbook_type(root):
                return next((node.get("ContentType") for node in root.findall(f"{{{namespace}}}Override")
                             if node.get("PartName") == "/xl/workbook.xml"), None)
            if workbook_type(old_types) != workbook_type(new_types):
                missing.append("[Content_Types].xml: workbook macro content type")
            if "xl/_rels/workbook.xml.rels" in old_parts and "xl/_rels/workbook.xml.rels" in new_parts:
                def vba_targets(archive):
                    root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
                    return sorted(node.get("Target") for node in root if (node.get("Type") or "").endswith("/vbaProject"))
                if vba_targets(before) != vba_targets(after):
                    missing.append("xl/_rels/workbook.xml.rels: VBA relationship")
        if missing:
            raise SaveTransactionError({
                "code": "EXCEL_SAVE_ADVANCED_PARTS_LOST",
                "message": "Staged workbook dropped advanced package content; source was not overwritten.",
                "parts": missing,
            })


def _requested_package_additions(session_data: dict, before, after) -> list[str]:
    relationship_base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    dirty = set(session_data.get("_dirty_paths") or [])
    before_snapshot = before.ensure_snapshot()
    after_snapshot = after.ensure_snapshot()
    old_sheets = {item["name"] for item in before_snapshot["workbook"]["sheets"]}
    new_sheets = {item["name"]: item for item in after_snapshot["workbook"]["sheets"]}
    expected_sheets = {
        sheet["target"] for name, sheet in new_sheets.items()
        if name not in old_sheets and f"sheets/{name}" in dirty
    }
    expected_links = Counter(
        (sheet["name"], value.get("target"))
        for sheet in session_data["sheets"]
        for cell, value in (sheet.get("hyperlinks") or {}).items()
        if f"sheets/{sheet['name']}/hyperlinks/{cell}" in dirty
        and isinstance(value, dict) and value.get("target")
    )
    expected_tables = {
        (sheet["name"], item.get("name"))
        for sheet in session_data["sheets"]
        for item in sheet.get("tables") or []
        if f"sheets/{sheet['name']}/tables/{item.get('name')}" in dirty
    }
    sheet_names_by_part = {item["target"]: name for name, item in new_sheets.items()}

    def related_sheet(source):
        prefix = "xl/worksheets/_rels/"
        if not source.startswith(prefix) or not source.endswith(".rels"):
            return None
        return sheet_names_by_part.get("xl/worksheets/" + source[len(prefix):-5])

    def only_expected_relationships(sheet_name, expected):
        old = (before_snapshot["worksheets"].get(sheet_name) or {}).get("relationships") or []
        new = (after_snapshot["worksheets"].get(sheet_name) or {}).get("relationships") or []
        old_records = Counter(json.dumps(item, sort_keys=True) for item in old)
        new_records = Counter(json.dumps(item, sort_keys=True) for item in new)
        additions = new_records - old_records
        return not (old_records - new_records) and additions == expected

    additions = []
    new_parts = set(after.parts) - set(before.parts)
    expected_sheet_relationships = {}
    prior = Counter((key.split("#", 1)[0], json.dumps(value, sort_keys=True))
                    for key, value in before.relationship_records.items())
    after_records = Counter((key.split("#", 1)[0], json.dumps(value, sort_keys=True))
                            for key, value in after.relationship_records.items())
    before_records = prior.copy()
    for key, record in after.relationship_records.items():
        identity = key.split("#", 1)[0], json.dumps(record, sort_keys=True)
        if prior[identity]:
            prior[identity] -= 1
            continue
        previous = before.relationship_records.get(key)
        if previous is not None and after_records[
            (key.split("#", 1)[0], json.dumps(previous, sort_keys=True))
        ] < before_records[(key.split("#", 1)[0], json.dumps(previous, sort_keys=True))]:
            continue
        source, _ = key.split("#", 1)
        target = record.get("target") or ""
        kind = record.get("type") or ""
        if (kind == relationship_base + "worksheet" and source == "xl/_rels/workbook.xml.rels"
                and record.get("target_mode") is None):
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            content_type = (after.content_type_records.get(f"Override:/{resolved}") or {}).get("ContentType")
            if (resolved in expected_sheets and resolved in new_parts
                    and f"Override:/{resolved}" not in before.content_type_records
                    and content_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"):
                additions.extend((f"package/relationships/{key}", f"package/content_types/Override:/{resolved}"))
        elif (kind == relationship_base + "hyperlink" and related_sheet(source)
              and record.get("target_mode") == "External"
              and expected_links[(related_sheet(source), target)] > 0):
            expected_links[(related_sheet(source), target)] -= 1
            additions.append(f"package/relationships/{key}")
            expected = {"type": record["type"], "target": target, "external": True}
            expected_sheet_relationships.setdefault(related_sheet(source), []).append(expected)
        elif (kind == relationship_base + "table" and expected_tables and related_sheet(source)
              and record.get("target_mode") is None):
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl/worksheets", target))
            if resolved in new_parts and resolved.startswith("xl/tables/"):
                root = ET.fromstring(after.parts[resolved])
                content_type = (after.content_type_records.get(f"Override:/{resolved}") or {}).get("ContentType")
                if ((related_sheet(source), root.get("name")) in expected_tables
                        and f"Override:/{resolved}" not in before.content_type_records
                        and content_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"):
                    additions.extend((f"package/relationships/{key}",
                                      f"package/content_types/Override:/{resolved}"))
                    expected = {"type": record["type"], "target": resolved, "external": False}
                    expected_sheet_relationships.setdefault(related_sheet(source), []).append(expected)
    for sheet_name, records in expected_sheet_relationships.items():
        expected = Counter(json.dumps(item, sort_keys=True) for item in records)
        if only_expected_relationships(sheet_name, expected):
            additions.append(f"worksheets/{sheet_name}/relationships")
    return additions


def _stage_row_heights(session_data: dict, reference_path: str | None, staging: Path) -> bool:
    dirty_paths = session_data.get("_dirty_paths") or []
    if (not reference_path or not dirty_paths
            or set(session_data.get("_dirty_features") or []) != {"row_properties"}
            or (session_data.get("_package_edits") or {}).get("upsert")
            or (session_data.get("_package_edits") or {}).get("delete")):
        return False
    row_paths = [re.fullmatch(r"sheets/([^/]+)/rows/(\d+)/height", path) for path in dirty_paths]
    if any(match is None for match in row_paths):
        return False

    rows_by_sheet = {sheet["name"]: sheet["rows"] for sheet in session_data["sheets"]}
    with zipfile.ZipFile(reference_path) as source:
        if any(name.startswith("_xmlsignatures/") for name in source.namelist()):
            raise SaveTransactionError({
                "code": "EXCEL_SAVE_REQUIRES_RESIGNING",
                "message": "Editing this signed workbook invalidates its package signature; no file was published.",
            })
        metadata = {name: source.read(name) for name in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels")}
        sheet_targets = {item["name"]: item["target"] for item in _workbook_snapshot(metadata)["sheets"]}
        updates: dict[str, dict[int, float | None]] = {}
        for match in row_paths:
            sheet_name, index_text = match.groups()
            row_index = int(index_text)
            if sheet_name not in rows_by_sheet or row_index >= len(rows_by_sheet[sheet_name]):
                raise SaveTransactionError({"code": "EXCEL_SAVE_ROW_NOT_FOUND", "message": "Edited row is absent from the session."})
            target = sheet_targets.get(sheet_name)
            if not target or target not in source.namelist():
                raise SaveTransactionError({"code": "EXCEL_SAVE_SHEET_NOT_FOUND", "message": "Edited worksheet is absent from the package."})
            height = rows_by_sheet[sheet_name][row_index].get("h")
            if height is not None and (not isinstance(height, (int, float)) or not math.isfinite(height)):
                raise ValueError("Row height must be a finite number or null.")
            updates.setdefault(target, {})[row_index + 1] = height

        row_tag = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?row(?=[\s/>])(?:\"[^\"]*\"|'[^']*'|[^'\">])*?>", re.DOTALL)
        row_number = re.compile(rb"(?<![\w:.-])r\s*=\s*([\"'])(\d+)\1")
        height_attrs = re.compile(rb"\s+(?:ht|customHeight)\s*=\s*(?:\"[^\"]*\"|'[^']*')")
        sheet_data_tag = re.compile(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?sheetData(?=[\s/>])(?:\"[^\"]*\"|'[^']*'|[^'\">])*?>", re.DOTALL)
        with zipfile.ZipFile(staging, "w") as output:
            output.comment = source.comment
            for info in source.infolist():
                raw = source.read(info.filename)
                pending = updates.get(info.filename)
                if pending:
                    sheet_data = sheet_data_tag.search(raw)
                    if sheet_data is None:
                        raise SaveTransactionError({"code": "EXCEL_SAVE_SHEET_DATA_NOT_FOUND", "message": "Worksheet has no sheetData element."})
                    prefix = sheet_data.group("prefix") or b""
                    closing = b"</" + prefix + b"sheetData>"
                    if sheet_data.group(0).endswith(b"/>"):
                        body_start = sheet_data.end()
                        body_end = body_start
                    else:
                        body_start = sheet_data.end()
                        body_end = raw.find(closing, body_start)
                        if body_end < 0:
                            raise SaveTransactionError({"code": "EXCEL_SAVE_SHEET_DATA_NOT_FOUND", "message": "Worksheet sheetData is not closed."})
                    body = raw[body_start:body_end]
                    seen = set()

                    def change_row(found):
                        tag = found.group(0)
                        number = row_number.search(tag)
                        index = int(number.group(2)) if number else None
                        if index not in pending:
                            return tag
                        if index in seen:
                            raise SaveTransactionError({"code": "EXCEL_SAVE_DUPLICATE_ROW", "message": "Worksheet has duplicate edited rows."})
                        seen.add(index)
                        tag = height_attrs.sub(b"", tag)
                        height = pending[index]
                        if height is None:
                            return tag
                        insertion = f' ht="{height}" customHeight="1"'.encode("ascii")
                        offset = -2 if tag.endswith(b"/>") else -1
                        return tag[:offset] + insertion + tag[offset:]

                    body = row_tag.sub(change_row, body)
                    for index in sorted(set(pending) - seen, reverse=True):
                        height = pending[index]
                        if height is None:
                            continue
                        new_row = (b"<" + prefix + b"row r=\"" + str(index).encode("ascii")
                                   + f'\" ht="{height}" customHeight="1"/>'.encode("ascii"))
                        insertion = len(body)
                        for found in row_tag.finditer(body):
                            number = row_number.search(found.group(0))
                            if number and int(number.group(2)) > index:
                                insertion = found.start()
                                break
                        body = body[:insertion] + new_row + body[insertion:]
                    if sheet_data.group(0).endswith(b"/>"):
                        raw = (raw[:sheet_data.start()] + sheet_data.group(0)[:-2] + b">"
                               + body + closing + raw[body_end:])
                    else:
                        raw = raw[:body_start] + body + raw[body_end:]
                output.writestr(info, raw)
    return True


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
    if reconstruct is reconstruct_excel and _stage_row_heights(to_write, verification_reference_path, staging):
        warnings = []
    else:
        warnings = list(reconstruct(to_write, str(staging)) or [])
    _guard_advanced_package_parts(
        verification_reference_path,
        staging,
        set((to_write.get("_package_edits") or {}).get("delete") or []),
    )

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
        if verify is verify_xlsx_preservation:
            requested_paths = list(requested_paths)
            for path in _requested_package_additions(session_data, left_inspection, right_inspection):
                if path not in requested_paths:
                    requested_paths.append(path)

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
            if intentional_edit and verification_reference_path:
                warnings.append("Preservation verification was not run for this edit; inspect the saved copy before relying on it.")
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
