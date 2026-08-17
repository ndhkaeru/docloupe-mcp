"""OOXML preservation verification and pre-save backup helpers."""
from __future__ import annotations

import copy
import fnmatch
import hashlib
import json
import os
import posixpath
import shutil
import tempfile
import threading
import time
import uuid
import warnings
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from functools import lru_cache
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_DOC_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_ID = f"{{{REL_DOC_NS}}}id"
BACKUP_RETENTION_DAYS = 2
BACKUP_PREFIX = "excel-backup-"
_CLEANUP_INTERVAL_SECONDS = 60 * 60
_CLEANUP_LOCK = threading.Lock()
_CLEANUP_STARTED = False
_TEXT_TAGS = {
    "t", "v", "f", "definedName", "formula1", "formula2", "text",
    "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter",
}
_ADVANCED_PREFIXES = (
    "xl/vbaProject", "xl/vmlDrawings/", "xl/pivot", "xl/slicer", "xl/externalLinks/",
    "xl/activeX/", "xl/embeddings/", "xl/threadedComments/", "xl/comments", "xl/persons/",
    "xl/customXml/", "customXml/", "customUI/", "_xmlsignatures/", "xl/drawings/",
    "xl/charts/", "xl/media/", "xl/printerSettings/", "xl/model/", "xl/connections",
)
_REQUIRED_PARTS = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"}


class _PackageParts(dict[str, bytes]):
    def __init__(self, values: dict[str, bytes] | None = None):
        super().__init__(values or {})
        self.xml_roots: dict[str, ET.Element] = {}
        self.invalid_xml: dict[str, str] = {}


@dataclass
class WorkbookInspection:
    path: str
    exists: bool
    size: int | None
    sha256: str | None
    state: tuple[int, int, int, int, int] | None
    parts: _PackageParts | None
    entry_names: tuple[str, ...]
    package_report: dict[str, Any]
    part_hashes: dict[str, str] = field(default_factory=dict)
    semantic_hashes: dict[str, str] = field(default_factory=dict)
    signature_hashes: dict[str, str] = field(default_factory=dict)
    relationship_records: dict[str, dict] = field(default_factory=dict)
    content_type_records: dict[str, dict] = field(default_factory=dict)
    loadable: bool = False
    load_error: str | None = None
    load_warnings: list[str] = field(default_factory=list)
    sheet_names: list[str] = field(default_factory=list)
    fingerprint_seconds: float = 0.0
    package_seconds: float = 0.0
    signature_seconds: float = 0.0
    load_probe_seconds: float = 0.0
    package_open_count: int = 0
    part_read_count: int = 0
    raw_part_bytes: int = 0
    reused_package_data: bool = False
    snapshot: dict[str, Any] | None = None
    snapshot_seconds: float = 0.0
    raw_parts_released: bool = False

    def file_metadata(self) -> dict[str, Any]:
        report = self.package_report
        return {
            "path": self.path,
            "exists": self.exists,
            "sha256": self.sha256,
            "size": self.size,
            "package_valid": bool(report.get("valid")),
            "package_errors": list(report.get("errors") or []),
            "package_warnings": list(report.get("warnings") or []),
            "part_count": report.get("part_count"),
            "features": copy.deepcopy(report.get("features") or {}),
            "loadable": self.loadable,
            "load_error": self.load_error,
            "load_warnings": list(self.load_warnings),
            "sheet_names": list(self.sheet_names),
        }

    def performance(self) -> dict[str, Any]:
        return {
            "fingerprint_seconds": round(self.fingerprint_seconds, 6),
            "package_seconds": round(self.package_seconds, 6),
            "signature_seconds": round(self.signature_seconds, 6),
            "load_probe_seconds": round(self.load_probe_seconds, 6),
            "snapshot_seconds": round(self.snapshot_seconds, 6),
            "package_open_count": self.package_open_count,
            "part_read_count": self.part_read_count,
            "raw_part_bytes": self.raw_part_bytes,
            "reused_package_data": self.reused_package_data,
            "snapshot_built": self.snapshot is not None,
            "raw_parts_released": self.raw_parts_released,
        }

    def ensure_snapshot(self) -> dict[str, Any]:
        if self.snapshot is None:
            if self.parts is None:
                raise RuntimeError("Raw package parts were released before snapshot creation.")
            started_at = time.perf_counter()
            self.snapshot = _package_snapshot_from_inspection(self)
            self.snapshot_seconds = time.perf_counter() - started_at
        return self.snapshot

    def release_raw_parts(self) -> None:
        self.parts = None
        self.raw_parts_released = True


class _MemorySampler:
    def __init__(self) -> None:
        self.before = _working_set_bytes()
        self.after = self.before
        self.peak = self.before
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        if self.before is None:
            return

        def sample() -> None:
            while not self._stop.wait(0.02):
                value = _working_set_bytes()
                if value is not None:
                    self.peak = max(self.peak or value, value)

        self._thread = threading.Thread(target=sample, daemon=True)
        self._thread.start()

    def finish(self) -> dict[str, int | None]:
        self._stop.set()
        if self._thread is not None:
            self._thread.join()
        self.after = _working_set_bytes()
        if self.after is not None:
            self.peak = max(self.peak or self.after, self.after)
        return {
            "memory_before_bytes": self.before,
            "memory_after_bytes": self.after,
            "peak_memory_bytes": self.peak,
        }


@lru_cache(maxsize=1)
def _windows_memory_api():
    if os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class ProcessMemoryCountersEx(ctypes.Structure):
            _fields_ = [
                ("cb", wintypes.DWORD),
                ("PageFaultCount", wintypes.DWORD),
                ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t),
                ("PrivateUsage", ctypes.c_size_t),
            ]

        psapi = ctypes.WinDLL("psapi")
        kernel32 = ctypes.WinDLL("kernel32")
        get_process_memory_info = psapi.GetProcessMemoryInfo
        get_process_memory_info.argtypes = (
            wintypes.HANDLE,
            ctypes.c_void_p,
            wintypes.DWORD,
        )
        get_process_memory_info.restype = wintypes.BOOL
        kernel32.GetCurrentProcess.restype = wintypes.HANDLE
        return (
            ctypes,
            ProcessMemoryCountersEx,
            get_process_memory_info,
            kernel32.GetCurrentProcess(),
        )
    except Exception:
        return None


def _working_set_bytes() -> int | None:
    if os.name != "nt":
        try:
            import resource

            usage = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
            return int(usage * (1024 if os.uname().sysname != "Darwin" else 1))
        except Exception:
            return None
    api = _windows_memory_api()
    if api is None:
        return None
    try:
        ctypes, counters_type, get_process_memory_info, process_handle = api
        counters = counters_type()
        counters.cb = ctypes.sizeof(counters)
        if not get_process_memory_info(
            process_handle,
            ctypes.byref(counters),
            counters.cb,
        ):
            return None
        return int(counters.WorkingSetSize)
    except Exception:
        return None


def _utc_now(now: datetime | None = None) -> datetime:
    value = now or datetime.now(timezone.utc)
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc)


def _iso(value: datetime) -> str:
    return value.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def _parse_iso(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(timezone.utc)


def _path_key(path: str | Path) -> str:
    return os.path.normcase(str(Path(path).expanduser().resolve()))


def _backup_directory() -> Path:
    configured = os.environ.get("DOCLOUPE_EXCEL_BACKUP_DIR")
    return Path(configured).expanduser() if configured else Path(tempfile.gettempdir()) / "docloupe-excel-backups"


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _file_sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _file_state(path: str | Path) -> tuple[int, int, int, int, int]:
    stat = Path(path).stat()
    return stat.st_dev, stat.st_ino, stat.st_size, stat.st_mtime_ns, stat.st_ctime_ns


def _file_fingerprint(path: str | Path) -> tuple[int, str, tuple[int, int, int, int, int]]:
    source = Path(path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"Excel file not found: {source}")
    before_state = _file_state(source)
    digest = _file_sha256(source)
    after_state = _file_state(source)
    if before_state != after_state:
        raise RuntimeError(f"Workbook changed while it was being fingerprinted: {path}")
    return after_state[2], digest, after_state


def _assert_file_unchanged(path: str | Path, expected_state: tuple[int, int, int, int, int]) -> None:
    if _file_state(path) != expected_state:
        raise RuntimeError(f"Workbook changed while it was being verified: {path}")


def _read_package_inventory(
    path: str | Path,
) -> tuple[_PackageParts, tuple[str, ...], int, int]:
    source = Path(path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"Excel file not found: {source}")
    try:
        with zipfile.ZipFile(source, "r") as archive:
            entries = tuple(item.filename for item in archive.infolist())
            values: dict[str, bytes] = {}
            for item in archive.infolist():
                if item.is_dir():
                    continue
                values[item.filename] = archive.read(item.filename)
    except zipfile.BadZipFile as exc:
        raise ValueError(f"Not a valid OOXML ZIP package: {source}") from exc
    parts = _PackageParts(values)
    for name, raw in parts.items():
        if name.endswith((".xml", ".rels")) or name == "[Content_Types].xml":
            try:
                parts.xml_roots[name] = ET.fromstring(raw)
            except Exception as exc:
                parts.invalid_xml[name] = str(exc)
    return parts, entries, 1, len(parts)


def _package_report_from_parts(
    path: str,
    parts: _PackageParts,
    entry_names: tuple[str, ...],
) -> dict[str, Any]:
    import re

    errors: list[str] = []
    package_warnings: list[str] = []
    names = set(parts)
    missing = sorted(_REQUIRED_PARTS - names)
    if missing:
        errors.append(f"missing required parts: {missing}")
    errors.extend(
        f"invalid XML in {name}: {message}"
        for name, message in sorted(parts.invalid_xml.items())
    )
    prefixes = {
        "drawings": "xl/drawings/",
        "charts": "xl/charts/",
        "media": "xl/media/",
        "vml_drawings": "xl/drawings/vmlDrawing",
        "pivot_tables": "xl/pivotTables/",
        "slicers": "xl/slicers/",
        "external_links": "xl/externalLinks/",
        "custom_xml": "customXml/",
    }
    features = {
        name: sum(1 for part in entry_names if part.startswith(prefix))
        for name, prefix in prefixes.items()
    }
    features["vba_project"] = int("xl/vbaProject.bin" in names)
    features["unknown_parts"] = sum(
        1
        for part in entry_names
        if not part.startswith(("_rels/", "docProps/", "xl/", "customXml/"))
        and part != "[Content_Types].xml"
    )

    referenced_parts: set[str] = set()
    for rel_part in sorted(name for name in parts if name.endswith(".rels")):
        rel_base = ""
        if rel_part == "_rels/.rels":
            rel_base = ""
        elif "/_rels/" in rel_part:
            folder, rel_name = rel_part.split("/_rels/", 1)
            rel_base = folder.rsplit("/", 1)[0] + "/" if "/" in folder else ""
            rel_base += rel_name[:-5]
        try:
            rel_xml = parts[rel_part].decode("utf-8")
            for match in re.finditer(r'\bTarget="([^"]+)"', rel_xml):
                target = match.group(1)
                if target.startswith(("http://", "https://", "mailto:")):
                    continue
                if target.startswith("/"):
                    referenced_parts.add(target.lstrip("/"))
                else:
                    referenced_parts.add(
                        posixpath.normpath(posixpath.join(posixpath.dirname(rel_base), target))
                    )
        except Exception:
            pass
    orphan_advanced = [
        name
        for name in parts
        if name.startswith(("xl/pivotTables/", "xl/externalLinks/"))
        and name not in referenced_parts
    ]
    if orphan_advanced:
        package_warnings.append(
            "advanced parts are present but not referenced by relationships: "
            f"{orphan_advanced[:8]}"
        )
    if any(
        features.get(name, 0)
        for name in (
            "vml_drawings",
            "pivot_tables",
            "slicers",
            "external_links",
            "vba_project",
        )
    ):
        package_warnings.append(
            "workbook contains advanced parts that are preserved best-effort only"
        )
    return {
        "path": path,
        "valid": not errors,
        "errors": errors,
        "warnings": package_warnings,
        "part_count": len(entry_names),
        "features": features,
    }


def _probe_workbook_load(path: str) -> tuple[bool, str | None, list[str], list[str], float]:
    workbook = None
    started_at = time.perf_counter()
    try:
        import openpyxl

        with warnings.catch_warnings(record=True) as captured:
            warnings.simplefilter("always")
            workbook = openpyxl.load_workbook(
                path,
                read_only=True,
                data_only=False,
                keep_vba=False,
                rich_text=True,
            )
            sheet_names = list(workbook.sheetnames)
        return (
            True,
            None,
            sorted({str(item.message) for item in captured}),
            sheet_names,
            time.perf_counter() - started_at,
        )
    except Exception as exc:
        return False, str(exc), [], [], time.perf_counter() - started_at
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass


def inspect_workbook(
    path: str | Path,
    *,
    check_loadable: bool = False,
    fingerprint: tuple[int, str, tuple[int, int, int, int, int]] | None = None,
    fingerprint_seconds: float | None = None,
) -> WorkbookInspection:
    resolved = str(Path(path).expanduser().resolve())
    source = Path(resolved)
    if fingerprint is None:
        fingerprint_started_at = time.perf_counter()
        try:
            fingerprint = _file_fingerprint(source)
        except Exception as exc:
            return WorkbookInspection(
                path=resolved,
                exists=source.is_file(),
                size=source.stat().st_size if source.is_file() else None,
                sha256=None,
                state=None,
                parts=None,
                entry_names=(),
                package_report={
                    "path": resolved,
                    "valid": False,
                    "errors": [str(exc)],
                    "warnings": [],
                    "part_count": None,
                    "features": {},
                },
                load_error=str(exc) if check_loadable else None,
                fingerprint_seconds=time.perf_counter() - fingerprint_started_at,
                raw_parts_released=True,
            )
        fingerprint_seconds = time.perf_counter() - fingerprint_started_at
    size, digest, state = fingerprint
    package_started_at = time.perf_counter()
    try:
        parts, entry_names, package_open_count, part_read_count = _read_package_inventory(source)
        package_error = None
    except Exception as exc:
        parts = None
        entry_names = ()
        package_open_count = 1
        part_read_count = 0
        package_error = str(exc)
    package_seconds = time.perf_counter() - package_started_at
    if parts is None:
        inspection = WorkbookInspection(
            path=resolved,
            exists=True,
            size=size,
            sha256=digest,
            state=state,
            parts=None,
            entry_names=entry_names,
            package_report={
                "path": resolved,
                "valid": False,
                "errors": [package_error or "Could not inspect workbook package."],
                "warnings": [],
                "part_count": None,
                "features": {},
            },
            fingerprint_seconds=float(fingerprint_seconds or 0.0),
            package_seconds=package_seconds,
            package_open_count=package_open_count,
            part_read_count=part_read_count,
            raw_parts_released=True,
        )
    else:
        part_hashes = {name: _sha256(raw) for name, raw in parts.items()}
        semantic_hashes = {
            name: _semantic_digest(name, raw, parts.xml_roots.get(name))
            for name, raw in parts.items()
        }
        signature_started_at = time.perf_counter()
        signature_hashes = {
            name: part_hashes[name]
            for name in sorted(part_hashes)
            if name.startswith("_xmlsignatures/")
        }
        signature_seconds = time.perf_counter() - signature_started_at
        try:
            relationship_records = _relationship_records(parts)
        except ValueError:
            relationship_records = {}
        try:
            content_type_records = _content_type_records(parts)
        except ValueError:
            content_type_records = {}
        inspection = WorkbookInspection(
            path=resolved,
            exists=True,
            size=size,
            sha256=digest,
            state=state,
            parts=parts,
            entry_names=entry_names,
            package_report=_package_report_from_parts(resolved, parts, entry_names),
            part_hashes=part_hashes,
            semantic_hashes=semantic_hashes,
            signature_hashes=signature_hashes,
            relationship_records=relationship_records,
            content_type_records=content_type_records,
            fingerprint_seconds=float(fingerprint_seconds or 0.0),
            package_seconds=package_seconds,
            signature_seconds=signature_seconds,
            package_open_count=package_open_count,
            part_read_count=part_read_count,
            raw_part_bytes=sum(len(raw) for raw in parts.values()),
        )
    if check_loadable and inspection.exists:
        (
            inspection.loadable,
            inspection.load_error,
            inspection.load_warnings,
            inspection.sheet_names,
            inspection.load_probe_seconds,
        ) = _probe_workbook_load(resolved)
    return inspection


def _clone_identical_inspection(
    source: WorkbookInspection,
    path: str,
    fingerprint: tuple[int, str, tuple[int, int, int, int, int]],
    fingerprint_seconds: float,
) -> WorkbookInspection:
    size, digest, state = fingerprint
    report = copy.deepcopy(source.package_report)
    report["path"] = path
    return WorkbookInspection(
        path=path,
        exists=True,
        size=size,
        sha256=digest,
        state=state,
        parts=source.parts,
        entry_names=source.entry_names,
        package_report=report,
        part_hashes=source.part_hashes,
        semantic_hashes=source.semantic_hashes,
        signature_hashes=source.signature_hashes,
        relationship_records=source.relationship_records,
        content_type_records=source.content_type_records,
        loadable=source.loadable,
        load_error=source.load_error,
        load_warnings=list(source.load_warnings),
        sheet_names=list(source.sheet_names),
        fingerprint_seconds=fingerprint_seconds,
        reused_package_data=True,
    )


def inspect_workbook_pair(
    before: str | Path,
    after: str | Path,
    *,
    check_loadable: bool = False,
) -> tuple[WorkbookInspection, WorkbookInspection]:
    before_path = str(Path(before).expanduser().resolve())
    after_path = str(Path(after).expanduser().resolve())
    left = inspect_workbook(before_path, check_loadable=check_loadable)
    if before_path == after_path:
        if left.size is None or left.sha256 is None or left.state is None:
            return left, copy.copy(left)
        return left, _clone_identical_inspection(
            left,
            after_path,
            (left.size, left.sha256, left.state),
            0.0,
        )
    fingerprint_started_at = time.perf_counter()
    try:
        right_fingerprint = _file_fingerprint(after_path)
    except Exception:
        return left, inspect_workbook(after_path, check_loadable=check_loadable)
    right_fingerprint_seconds = time.perf_counter() - fingerprint_started_at
    if (
        left.size == right_fingerprint[0]
        and left.sha256 is not None
        and left.sha256 == right_fingerprint[1]
    ):
        return left, _clone_identical_inspection(
            left,
            after_path,
            right_fingerprint,
            right_fingerprint_seconds,
        )
    right = inspect_workbook(
        after_path,
        check_loadable=check_loadable,
        fingerprint=right_fingerprint,
        fingerprint_seconds=right_fingerprint_seconds,
    )
    return left, right


def inspection_pair_performance(
    before: WorkbookInspection,
    after: WorkbookInspection,
    *,
    metadata_seconds: float | None = None,
    semantic_verification_seconds: float | None = None,
    total_tool_seconds: float | None = None,
    memory: dict[str, int | None] | None = None,
) -> dict[str, Any]:
    memory = memory or {}
    return {
        "metadata_seconds": round(float(metadata_seconds or 0.0), 6),
        "signature_seconds": round(
            before.signature_seconds + after.signature_seconds,
            6,
        ),
        "semantic_verification_seconds": round(
            float(semantic_verification_seconds or 0.0),
            6,
        ),
        "total_tool_seconds": round(float(total_tool_seconds or 0.0), 6),
        "package_open_count": before.package_open_count + after.package_open_count,
        "part_read_count": before.part_read_count + after.part_read_count,
        "raw_part_bytes": before.raw_part_bytes + after.raw_part_bytes,
        "memory_before_bytes": memory.get("memory_before_bytes"),
        "memory_after_bytes": memory.get("memory_after_bytes"),
        "peak_memory_bytes": memory.get("peak_memory_bytes"),
        "before": before.performance(),
        "after": after.performance(),
    }


def _read_package(path: str | Path) -> dict[str, bytes]:
    inspection = inspect_workbook(path)
    if inspection.parts is None:
        errors = inspection.package_report.get("errors") or ["Could not read workbook package."]
        raise ValueError(str(errors[0]))
    return inspection.parts


def _signature_hashes_only(path: str | Path | None) -> tuple[dict[str, str], dict[str, Any]]:
    if path is None:
        return {}, {
            "mode": "signature_only",
            "package_open_count": 0,
            "part_read_count": 0,
            "artifact_bytes": 0,
            "elapsed_seconds": 0.0,
        }
    source = Path(path).expanduser().resolve()
    started_at = time.perf_counter()
    try:
        with zipfile.ZipFile(source, "r") as archive:
            names = sorted(
                name
                for name in archive.namelist()
                if name.startswith("_xmlsignatures/") and not name.endswith("/")
            )
            values = {name: archive.read(name) for name in names}
    except zipfile.BadZipFile as exc:
        raise ValueError(f"Not a valid OOXML ZIP package: {source}") from exc
    return (
        {name: _sha256(raw) for name, raw in values.items()},
        {
            "mode": "signature_only",
            "package_open_count": 1,
            "part_read_count": len(values),
            "artifact_bytes": sum(len(raw) for raw in values.values()),
            "elapsed_seconds": round(time.perf_counter() - started_at, 6),
        },
    )


def _package_signature_hashes(
    path: str | Path | None,
    inspection: WorkbookInspection | None = None,
) -> tuple[dict[str, str], dict[str, Any]]:
    if inspection is not None:
        return dict(inspection.signature_hashes), {
            "mode": "reused_full_inspection",
            "package_open_count": 0,
            "part_read_count": 0,
            "artifact_bytes": 0,
            "elapsed_seconds": round(inspection.signature_seconds, 6),
        }
    return _signature_hashes_only(path)


def package_signature_report(
    before_path: str | Path | None,
    after_path: str | Path,
    intentional_edit: bool,
    *,
    before_inspection: WorkbookInspection | None = None,
    after_inspection: WorkbookInspection | None = None,
) -> dict:
    """Report byte preservation and conservative invalidation status for OOXML signatures."""
    started_at = time.perf_counter()
    before_parts, before_performance = _package_signature_hashes(
        before_path,
        before_inspection,
    )
    after_parts, after_performance = _package_signature_hashes(
        after_path,
        after_inspection,
    )
    present = bool(before_parts or after_parts)
    parts_preserved = before_parts == after_parts
    intentional_edit = bool(intentional_edit)

    if not present:
        status = "not_present"
        message = "No OOXML package signature parts are present."
    elif not parts_preserved:
        status = "signature_parts_changed"
        message = (
            "CRITICAL: OOXML package signature parts were added, removed, or modified during save. "
            "Treat existing package signatures as invalid until the workbook is signed again."
        )
    elif intentional_edit:
        status = "requires_resigning"
        message = (
            "Workbook content was intentionally edited while OOXML package signature parts were "
            "byte-preserved. Treat existing signatures as invalid or unverified until the workbook "
            "is signed again."
        )
    else:
        status = "preserved_unverified"
        message = (
            "OOXML package signature parts were byte-preserved. Cryptographic signature validity "
            "was not verified."
        )

    return {
        "present": present,
        "intentional_edit": intentional_edit,
        "status": status,
        "parts_before": before_parts,
        "parts_after": after_parts,
        "parts_preserved": parts_preserved,
        "message": message,
        "performance": {
            "before": before_performance,
            "after": after_performance,
            "total_seconds": round(time.perf_counter() - started_at, 6),
        },
    }


def _local(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def _xml_node(element: ET.Element | None):
    if element is None:
        return None
    attrs = {key: value for key, value in sorted(element.attrib.items())}
    local = _local(element.tag)
    text = element.text or ""
    if local not in _TEXT_TAGS:
        text = text.strip()
    result = {"tag": element.tag}
    if attrs:
        result["attrs"] = attrs
    if text:
        result["text"] = text
    children = [_xml_node(child) for child in list(element)]
    if children:
        result["children"] = children
    return result


def _xml_root(parts: dict[str, bytes], part_name: str) -> ET.Element | None:
    if isinstance(parts, _PackageParts):
        if part_name in parts.invalid_xml:
            raise ValueError(f"Invalid XML part {part_name}: {parts.invalid_xml[part_name]}")
        cached = parts.xml_roots.get(part_name)
        if cached is not None:
            return cached
    raw = parts.get(part_name)
    if raw is None:
        return None
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise ValueError(f"Invalid XML part {part_name}: {exc}") from exc
    if isinstance(parts, _PackageParts):
        parts.xml_roots[part_name] = root
    return root


def _child(parent: ET.Element | None, name: str) -> ET.Element | None:
    if parent is None:
        return None
    return next((item for item in list(parent) if _local(item.tag) == name), None)


def _children(parent: ET.Element | None, name: str) -> list[ET.Element]:
    if parent is None:
        return []
    return [item for item in list(parent) if _local(item.tag) == name]


def _descendants(parent: ET.Element | None, name: str) -> list[ET.Element]:
    if parent is None:
        return []
    return [item for item in parent.iter() if _local(item.tag) == name]


def _relationship_part(source_part: str) -> str:
    directory, filename = posixpath.split(source_part)
    return posixpath.join(directory, "_rels", filename + ".rels")


def _resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _relationships(parts: dict[str, bytes], source_part: str) -> dict[str, dict]:
    root = _xml_root(parts, _relationship_part(source_part))
    relationships: dict[str, dict] = {}
    for rel in list(root) if root is not None else []:
        rel_id = rel.attrib.get("Id")
        if not rel_id:
            continue
        target = rel.attrib.get("Target", "")
        external = rel.attrib.get("TargetMode") == "External"
        relationships[rel_id] = {
            "type": rel.attrib.get("Type"),
            "target": target if external else _resolve_target(source_part, target),
            "external": external,
        }
    return relationships


def _semantic_digest(
    part_name: str,
    raw: bytes,
    root: ET.Element | None = None,
) -> str:
    if part_name.endswith((".xml", ".rels")) or part_name == "[Content_Types].xml":
        try:
            value = _xml_node(root if root is not None else ET.fromstring(raw))
            raw = json.dumps(
                value,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
        except ET.ParseError:
            pass
    return _sha256(raw)


def _relationship_records(parts: dict[str, bytes]) -> dict[str, dict]:
    records: dict[str, dict] = {}
    for part_name in sorted(name for name in parts if name.endswith(".rels")):
        root = _xml_root(parts, part_name)
        if root is None:
            continue
        for index, rel in enumerate(list(root)):
            rel_id = rel.attrib.get("Id") or f"index:{index}"
            records[f"{part_name}#{rel_id}"] = {
                "type": rel.attrib.get("Type"),
                "target": rel.attrib.get("Target"),
                "target_mode": rel.attrib.get("TargetMode"),
            }
    return records


def _content_type_records(parts: dict[str, bytes]) -> dict[str, dict]:
    root = _xml_root(parts, "[Content_Types].xml")
    records: dict[str, dict] = {}
    for index, item in enumerate(list(root) if root is not None else []):
        local = _local(item.tag)
        identity = item.attrib.get("PartName") or item.attrib.get("Extension") or f"index:{index}"
        records[f"{local}:{identity}"] = dict(sorted(item.attrib.items()))
    return records


def _mapping_diff(before: dict, after: dict) -> dict:
    before_keys = set(before)
    after_keys = set(after)
    return {
        "added": {key: after[key] for key in sorted(after_keys - before_keys)},
        "removed": {key: before[key] for key in sorted(before_keys - after_keys)},
        "modified": {
            key: {"before": before[key], "after": after[key]}
            for key in sorted(before_keys & after_keys)
            if before[key] != after[key]
        },
    }


def _package_content_diff_from_parts(
    before_parts: dict[str, bytes],
    after_parts: dict[str, bytes],
) -> dict:
    before_names = set(before_parts)
    after_names = set(after_parts)
    common = before_names & after_names
    modified = sorted(
        name for name in common
        if _semantic_digest(name, before_parts[name]) != _semantic_digest(name, after_parts[name])
    )
    added = sorted(after_names - before_names)
    removed = sorted(before_names - after_names)
    modified_xml = [
        name for name in modified
        if name.endswith((".xml", ".rels")) or name == "[Content_Types].xml"
    ]
    modified_binary = [name for name in modified if name not in modified_xml]
    relationship_diff = _mapping_diff(
        _relationship_records(before_parts),
        _relationship_records(after_parts),
    )
    content_type_diff = _mapping_diff(
        _content_type_records(before_parts),
        _content_type_records(after_parts),
    )
    binary_hash_changes = [
        {
            "part": name,
            "before_sha256": _sha256(before_parts[name]),
            "after_sha256": _sha256(after_parts[name]),
        }
        for name in modified_binary
    ]
    changed = bool(added or removed or modified)
    return {
        "before_part_count": len(before_names),
        "after_part_count": len(after_names),
        "added": added,
        "removed": removed,
        "modified": modified,
        "modified_xml": modified_xml,
        "modified_binary": modified_binary,
        "binary_hash_changes": binary_hash_changes,
        "relationship_changes": relationship_diff,
        "content_type_changes": content_type_diff,
        "changed": changed,
    }


def _package_content_diff_from_inspections(
    before: WorkbookInspection,
    after: WorkbookInspection,
) -> dict:
    before_names = set(before.part_hashes)
    after_names = set(after.part_hashes)
    common = before_names & after_names
    modified = sorted(
        name
        for name in common
        if before.semantic_hashes.get(name) != after.semantic_hashes.get(name)
    )
    added = sorted(after_names - before_names)
    removed = sorted(before_names - after_names)
    modified_xml = [
        name
        for name in modified
        if name.endswith((".xml", ".rels")) or name == "[Content_Types].xml"
    ]
    modified_binary = [name for name in modified if name not in modified_xml]
    relationship_diff = _mapping_diff(
        before.relationship_records,
        after.relationship_records,
    )
    content_type_diff = _mapping_diff(
        before.content_type_records,
        after.content_type_records,
    )
    binary_hash_changes = [
        {
            "part": name,
            "before_sha256": before.part_hashes[name],
            "after_sha256": after.part_hashes[name],
        }
        for name in modified_binary
    ]
    return {
        "before_part_count": len(before_names),
        "after_part_count": len(after_names),
        "added": added,
        "removed": removed,
        "modified": modified,
        "modified_xml": modified_xml,
        "modified_binary": modified_binary,
        "binary_hash_changes": binary_hash_changes,
        "relationship_changes": relationship_diff,
        "content_type_changes": content_type_diff,
        "changed": bool(added or removed or modified),
    }


def package_content_diff(before: str, after: str) -> dict:
    """Compare package parts, relationships, content types, and semantic hashes."""
    left, right = inspect_workbook_pair(before, after)
    try:
        return _package_content_diff_from_inspections(left, right)
    finally:
        left.release_raw_parts()
        right.release_raw_parts()


def _validate_package_xml(parts: dict[str, bytes]) -> None:
    if isinstance(parts, _PackageParts) and parts.invalid_xml:
        name = sorted(parts.invalid_xml)[0]
        raise ValueError(f"Invalid XML package part: {name}")
    for name, raw in parts.items():
        if name.endswith((".xml", ".rels")) or name == "[Content_Types].xml":
            if isinstance(parts, _PackageParts) and name in parts.xml_roots:
                continue
            try:
                ET.fromstring(raw)
            except ET.ParseError as exc:
                raise ValueError(f"Invalid XML package part: {name}") from exc


def _byte_identical_part_diff(parts: dict[str, bytes]) -> dict:
    part_count = len(parts)
    return {
        "before_part_count": part_count,
        "after_part_count": part_count,
        "added": [],
        "removed": [],
        "modified": [],
        "modified_xml": [],
        "modified_binary": [],
        "binary_hash_changes": [],
        "relationship_changes": _mapping_diff({}, {}),
        "content_type_changes": _mapping_diff({}, {}),
        "changed": False,
    }


def _rich_text(container: ET.Element | None) -> dict | None:
    if container is None:
        return None
    runs = []
    offset = 0
    direct_runs = _children(container, "r")
    source_runs = direct_runs or [container]
    for run in source_runs:
        text_node = _child(run, "t")
        text = text_node.text if text_node is not None and text_node.text is not None else ""
        runs.append({
            "text": text,
            "start": offset,
            "end": offset + len(text),
            "text_attrs": dict(sorted(text_node.attrib.items())) if text_node is not None and text_node.attrib else {},
            "properties": _xml_node(_child(run, "rPr")) if direct_runs else None,
        })
        offset += len(text)
    phonetic_runs = [_xml_node(item) for item in list(container) if _local(item.tag) == "rPh"]
    phonetic_properties = _xml_node(next(
        (item for item in list(container) if _local(item.tag) == "phoneticPr"),
        None,
    ))
    return {
        "text": "".join(run["text"] for run in runs),
        "run_count": len(runs),
        "runs": runs,
        "phonetic_runs": phonetic_runs,
        "phonetic_properties": phonetic_properties,
    }


def _shared_strings(parts: dict[str, bytes]) -> list[dict]:
    root = _xml_root(parts, "xl/sharedStrings.xml")
    return [_rich_text(item) for item in _children(root, "si")] if root is not None else []


def _style_snapshot(parts: dict[str, bytes]) -> dict:
    root = _xml_root(parts, "xl/styles.xml")
    if root is None:
        return {"cell_xfs": [], "named_styles": [], "dxfs": None, "table_styles": None, "colors": None}

    def nodes(section: str, child_name: str) -> list[ET.Element]:
        return _children(_child(root, section), child_name)

    numfmts = {item.attrib.get("numFmtId"): item.attrib.get("formatCode") for item in nodes("numFmts", "numFmt")}
    fonts = [_xml_node(item) for item in nodes("fonts", "font")]
    fills = [_xml_node(item) for item in nodes("fills", "fill")]
    borders = [_xml_node(item) for item in nodes("borders", "border")]
    style_xfs = nodes("cellStyleXfs", "xf")
    cell_xfs = nodes("cellXfs", "xf")
    style_xf_snapshots = [_xml_node(item) for item in style_xfs]

    def indexed(values: list, raw: str | None):
        try:
            index = int(raw or 0)
        except ValueError:
            return {"invalid_index": raw}
        return values[index] if 0 <= index < len(values) else {"missing_index": index}

    def resolved_xf(xf: ET.Element) -> dict:
        reference_attrs = {"numFmtId", "fontId", "fillId", "borderId", "xfId"}
        attrs = {key: value for key, value in sorted(xf.attrib.items()) if _local(key) not in reference_attrs}
        numfmt_id = xf.attrib.get("numFmtId", "0")
        try:
            numfmt_key = str(int(numfmt_id))
        except ValueError:
            numfmt_key = numfmt_id
        return {
            "attrs": attrs,
            "numfmt": numfmts.get(numfmt_key, f"builtin:{numfmt_key}"),
            "font": indexed(fonts, xf.attrib.get("fontId")),
            "fill": indexed(fills, xf.attrib.get("fillId")),
            "border": indexed(borders, xf.attrib.get("borderId")),
            "base_style": indexed(style_xf_snapshots, xf.attrib.get("xfId")),
            "alignment": _xml_node(_child(xf, "alignment")),
            "protection": _xml_node(_child(xf, "protection")),
            "extensions": _xml_node(_child(xf, "extLst")),
        }

    resolved_style_xfs = [resolved_xf(item) for item in style_xfs]
    resolved = [resolved_xf(item) for item in cell_xfs]
    named_styles = []
    for item in nodes("cellStyles", "cellStyle"):
        style_id = item.attrib.get("xfId")
        entry = {"attrs": {key: value for key, value in sorted(item.attrib.items()) if _local(key) != "xfId"}}
        entry["style"] = indexed(resolved_style_xfs, style_id)
        named_styles.append(entry)
    named_styles.sort(key=lambda value: json.dumps(value, sort_keys=True, ensure_ascii=False))
    return {
        "cell_xfs": resolved,
        "named_styles": named_styles,
        "dxfs": _xml_node(_child(root, "dxfs")),
        "table_styles": _xml_node(_child(root, "tableStyles")),
        "colors": _xml_node(_child(root, "colors")),
    }


def _workbook_snapshot(parts: dict[str, bytes]) -> dict:
    root = _xml_root(parts, "xl/workbook.xml")
    if root is None:
        return {"missing": True, "sheets": []}
    rels = _relationships(parts, "xl/workbook.xml")
    sheets = []
    for item in _children(_child(root, "sheets"), "sheet"):
        rel = rels.get(item.attrib.get(REL_ID, ""), {})
        sheets.append({
            "name": item.attrib.get("name"),
            "sheet_id": item.attrib.get("sheetId"),
            "state": item.attrib.get("state", "visible"),
            "target": rel.get("target"),
        })
    defined_names = [_xml_node(item) for item in _children(_child(root, "definedNames"), "definedName")]
    defined_names.sort(key=lambda value: json.dumps(value, sort_keys=True, ensure_ascii=False))
    return {
        "workbook_properties": _xml_node(_child(root, "workbookPr")),
        "workbook_protection": _xml_node(_child(root, "workbookProtection")),
        "calculation": _xml_node(_child(root, "calcPr")),
        "views": _xml_node(_child(root, "bookViews")),
        "sheets": sheets,
        "defined_names": defined_names,
        "external_references": _xml_node(_child(root, "externalReferences")),
        "extensions": _xml_node(_child(root, "extLst")),
    }


def _cell_snapshot(
    cell: ET.Element,
    shared_strings: list[dict],
    styles: list[dict],
) -> dict:
    cell_type = cell.attrib.get("t")
    style_id = cell.attrib.get("s", "0")
    try:
        style_index = int(style_id)
    except ValueError:
        style_index = -1
    style = styles[style_index] if 0 <= style_index < len(styles) else {"missing_style": style_id}
    formula_node = _child(cell, "f")
    value_node = _child(cell, "v")
    inline = _child(cell, "is")
    rich = None
    logical_value = value_node.text if value_node is not None else None
    if cell_type == "s" and logical_value is not None:
        try:
            shared_string_index = int(logical_value)
            rich = dict(shared_strings[shared_string_index])
            rich["storage_form"] = "sharedStrings"
            rich["shared_string_index"] = shared_string_index
            logical_value = rich["text"]
        except (ValueError, IndexError):
            logical_value = {"invalid_shared_string": logical_value}
    elif cell_type == "inlineStr":
        rich = _rich_text(inline)
        if rich is not None:
            rich = dict(rich)
            rich["storage_form"] = "inlineStr"
            rich["shared_string_index"] = None
        logical_value = rich["text"] if rich else ""
    cache_state = "missing" if value_node is None else ("empty" if value_node.text in (None, "") else "value")
    if formula_node is not None:
        presence_kind = "formula"
    elif inline is not None:
        presence_kind = "inline_string"
    elif value_node is not None:
        presence_kind = "empty_value" if value_node.text in (None, "") else "value"
    else:
        presence_kind = "explicit_empty"
    return {
        "present": True,
        "presence_kind": presence_kind,
        "type": cell_type,
        "value": logical_value,
        "formula": {
            "text": formula_node.text or "",
            "attrs": dict(sorted(formula_node.attrib.items())),
            "cached_value": value_node.text if value_node is not None else None,
            "cache_state": cache_state,
        } if formula_node is not None else None,
        "rich_text": rich,
        "style": style,
    }


def _relationship_inventory(rels: dict[str, dict]) -> list[dict]:
    values = list(rels.values())
    values.sort(key=lambda value: json.dumps(value, sort_keys=True, ensure_ascii=False))
    return values


def _related_xml(parts: dict[str, bytes], rels: dict[str, dict], suffix: str) -> list[dict]:
    values = []
    for rel in rels.values():
        if not (rel.get("type") or "").endswith(suffix) or rel.get("external"):
            continue
        target = rel.get("target")
        values.append({"target": target, "content": _xml_node(_xml_root(parts, target)) if target in parts else {"missing": True}})
    values.sort(key=lambda value: value.get("target") or "")
    return values


def _worksheet_snapshot(
    parts: dict[str, bytes],
    target: str,
    shared_strings: list[dict],
    styles: list[dict],
) -> dict:
    root = _xml_root(parts, target)
    if root is None:
        return {"missing": True, "cells": {}, "rows": {}}
    rels = _relationships(parts, target)
    rows: dict[str, dict] = {}
    cells: dict[str, dict] = {}
    sheet_data = _child(root, "sheetData")
    for row in _children(sheet_data, "row"):
        row_number = row.attrib.get("r", str(len(rows) + 1))
        attrs = {key: value for key, value in sorted(row.attrib.items()) if _local(key) != "r"}
        if attrs:
            rows[row_number] = attrs
        for cell in _children(row, "c"):
            reference = cell.attrib.get("r", f"row:{row_number}:cell:{len(cells)}")
            cells[reference] = _cell_snapshot(cell, shared_strings, styles)

    hyperlinks = []
    for item in _children(_child(root, "hyperlinks"), "hyperlink"):
        attrs = dict(sorted(item.attrib.items()))
        rel_id = attrs.pop(REL_ID, None)
        entry = {"attrs": attrs}
        if rel_id:
            entry["relationship"] = rels.get(rel_id, {"missing_relationship": rel_id})
        hyperlinks.append(entry)
    hyperlinks.sort(key=lambda value: json.dumps(value, sort_keys=True, ensure_ascii=False))

    return {
        "sheet_properties": _xml_node(_child(root, "sheetPr")),
        "sheet_views": _xml_node(_child(root, "sheetViews")),
        "row_defaults": _xml_node(_child(root, "sheetFormatPr")),
        "columns": _xml_node(_child(root, "cols")),
        "merged_cells": _xml_node(_child(root, "mergeCells")),
        "printing": {
            name: _xml_node(_child(root, name))
            for name in ("printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks")
        },
        "filtering": {
            "auto_filter": _xml_node(_child(root, "autoFilter")),
            "sort_state": _xml_node(_child(root, "sortState")),
        },
        "validation": {
            "data_validations": _xml_node(_child(root, "dataValidations")),
            "conditional_formatting": [_xml_node(item) for item in _children(root, "conditionalFormatting")],
        },
        "protection": {
            "sheet": _xml_node(_child(root, "sheetProtection")),
            "ranges": _xml_node(_child(root, "protectedRanges")),
        },
        "ignored_errors": _xml_node(_child(root, "ignoredErrors")),
        "hyperlinks": hyperlinks,
        "rows": rows,
        "cells": cells,
        "tables": _related_xml(parts, rels, "/table"),
        "comments": _related_xml(parts, rels, "/comments"),
        "relationships": _relationship_inventory(rels),
        "extensions": _xml_node(_child(root, "extLst")),
    }


def _document_properties(parts: dict[str, bytes]) -> dict:
    result = {}
    for name in ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"):
        result[name] = _xml_node(_xml_root(parts, name)) if name in parts else None
    return result


def _advanced_parts(
    parts: dict[str, bytes],
    semantic_hashes: dict[str, str] | None = None,
) -> dict[str, str]:
    return {
        name: (
            semantic_hashes[name]
            if semantic_hashes is not None and name in semantic_hashes
            else _semantic_digest(name, raw)
        )
        for name, raw in parts.items()
        if name == "xl/calcChain.xml" or name.startswith(_ADVANCED_PREFIXES)
    }


def _package_snapshot_from_inspection(inspection: WorkbookInspection) -> dict[str, Any]:
    parts = inspection.parts
    if parts is None:
        raise RuntimeError("Raw package parts are unavailable for semantic snapshot creation.")
    workbook = _workbook_snapshot(parts)
    styles = _style_snapshot(parts)
    shared_strings = _shared_strings(parts)
    worksheets = {}
    for sheet in workbook.get("sheets", []):
        target = sheet.get("target")
        if target:
            worksheets[sheet.get("name") or target] = _worksheet_snapshot(
                parts,
                target,
                shared_strings,
                styles["cell_xfs"],
            )
    return {
        "workbook": workbook,
        "styles": styles,
        "worksheets": worksheets,
        "document_properties": _document_properties(parts),
        "theme": inspection.semantic_hashes.get("xl/theme/theme1.xml"),
        "advanced_parts": _advanced_parts(parts, inspection.semantic_hashes),
    }


def _package_snapshot(path: str) -> dict:
    inspection = inspect_workbook(path)
    snapshot = inspection.ensure_snapshot()
    return {**snapshot, "parts": inspection.parts}


def _preview(value):
    encoded = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    if len(encoded) <= 1600:
        return value
    return {
        "summary": encoded[:1400] + "…",
        "sha256": _sha256(encoded.encode("utf-8")),
        "serialized_length": len(encoded),
    }


def _matches_path(path: str, patterns: tuple[str, ...]) -> bool:
    return any(path == pattern or fnmatch.fnmatchcase(path, pattern) for pattern in patterns)


def _normalize_approved_normalization_rules(
    items: list[str | dict] | tuple[str | dict, ...] | None,
) -> tuple[dict, ...]:
    rules = []
    for index, item in enumerate(items or ()):
        if isinstance(item, str):
            path = item.strip()
            if not path:
                raise ValueError(f"approved_normalizations[{index}] requires a non-empty path.")
            rules.append({
                "path": path,
                "rationale": None,
                "bidirectional": False,
                "evidence": None,
                "complete": False,
                "issues": ["rationale", "bidirectional", "evidence"],
            })
            continue
        if not isinstance(item, dict):
            raise ValueError(
                f"approved_normalizations[{index}] must be a path string or rule object."
            )
        path = str(item.get("path") or item.get("pattern") or "").strip()
        if not path:
            raise ValueError(f"approved_normalizations[{index}] requires a non-empty path.")
        rationale = str(item.get("rationale") or "").strip() or None
        bidirectional = item.get("bidirectional") is True
        evidence = item.get("evidence")
        evidence_complete = (
            isinstance(evidence, dict)
            and "before" in evidence
            and "after" in evidence
        )
        issues = []
        if rationale is None:
            issues.append("rationale")
        if not bidirectional:
            issues.append("bidirectional")
        if not evidence_complete:
            issues.append("evidence")
        rules.append({
            "path": path,
            "rationale": rationale,
            "bidirectional": bidirectional,
            "evidence": evidence if evidence_complete else None,
            "complete": not issues,
            "issues": issues,
        })
    return tuple(rules)


def _normalization_classification(
    path: str,
    before,
    after,
    rules: tuple[dict, ...],
) -> str | None:
    matching = [rule for rule in rules if _matches_path(path, (rule["path"],))]
    if not matching:
        return None
    for rule in matching:
        if not rule["complete"]:
            continue
        evidence = rule["evidence"]
        forward = before == evidence["before"] and after == evidence["after"]
        reverse = before == evidence["after"] and after == evidence["before"]
        if forward or (rule["bidirectional"] and reverse):
            return "APPROVED_NORMALIZATION"
    return "VERIFIER_GAP"


class _Changes:
    def __init__(
        self,
        limit: int,
        requested_paths: tuple[str, ...] = (),
        approved_normalization_rules: tuple[dict, ...] = (),
        fixture_gap_paths: tuple[str, ...] = (),
        verifier_gap_paths: tuple[str, ...] = (),
    ):
        self.limit = limit
        self.requested_paths = requested_paths
        self.approved_normalization_rules = approved_normalization_rules
        self.fixture_gap_paths = fixture_gap_paths
        self.verifier_gap_paths = verifier_gap_paths
        self.total = 0
        self.items: list[dict] = []
        self.severity_counts = {"critical": 0, "high": 0, "medium": 0, "info": 0}
        self.classification_counts = {
            "REQUESTED": 0,
            "APPROVED_NORMALIZATION": 0,
            "UNAPPROVED_LOSS": 0,
            "FIXTURE_GAP": 0,
            "VERIFIER_GAP": 0,
            "PACKAGE_INVALID": 0,
        }

    def _classification(self, path: str, before, after) -> str:
        if _matches_path(path, self.requested_paths):
            return "REQUESTED"
        normalization = _normalization_classification(
            path,
            before,
            after,
            self.approved_normalization_rules,
        )
        if normalization:
            return normalization
        if _matches_path(path, self.fixture_gap_paths):
            return "FIXTURE_GAP"
        if _matches_path(path, self.verifier_gap_paths):
            return "VERIFIER_GAP"
        return "UNAPPROVED_LOSS"

    def add(self, path: str, category: str, severity: str, before, after) -> None:
        if before == after:
            return
        classification = self._classification(path, before, after)
        self.total += 1
        self.severity_counts[severity] = self.severity_counts.get(severity, 0) + 1
        self.classification_counts[classification] += 1
        if len(self.items) < self.limit:
            self.items.append({
                "path": path,
                "category": category,
                "severity": severity,
                "classification": classification,
                "before": _preview(before),
                "after": _preview(after),
            })


def _compare_mapping(
    changes: _Changes,
    base_path: str,
    category: str,
    severity: str,
    before: dict,
    after: dict,
) -> None:
    for key in sorted(set(before) | set(after)):
        changes.add(f"{base_path}/{key}", category, severity, before.get(key), after.get(key))


def _compare_cells(changes: _Changes, sheet_name: str, before: dict, after: dict) -> None:
    all_refs = sorted(set(before) | set(after))
    for reference in all_refs:
        left = before.get(reference)
        right = after.get(reference)
        path = f"worksheets/{sheet_name}/cells/{reference}"
        if left is None or right is None:
            important = (left or right or {}).get("formula") or (left or right or {}).get("rich_text") or (
                (left or right or {}).get("style") not in (None, {}, {"missing_style": "0"})
            )
            changes.add(path, "cell_presence", "high" if important else "info", left, right)
            continue
        changes.add(path + "/value", "cell_value", "info", left.get("value"), right.get("value"))
        changes.add(path + "/type", "cell_type", "medium", left.get("type"), right.get("type"))
        changes.add(path + "/formula", "formula", "high", left.get("formula"), right.get("formula"))
        left_rich = left.get("rich_text")
        right_rich = right.get("rich_text")
        rich_is_material = any(
            rich and (len(rich.get("runs", [])) > 1 or rich.get("phonetic_runs") or rich.get("phonetic_properties") or any(run.get("properties") for run in rich.get("runs", [])))
            for rich in (left_rich, right_rich)
        )
        if rich_is_material:
            changes.add(path + "/rich_text", "rich_text", "high", left_rich, right_rich)
        changes.add(path + "/style", "cell_style", "high", left.get("style"), right.get("style"))


def _verification_report(
    before_path: str,
    after_path: str,
    before_sha256: str,
    after_sha256: str,
    before_size: int,
    after_size: int,
    requested: tuple[str, ...],
    approved_rules: tuple[dict, ...],
    fixture_gaps: tuple[str, ...],
    verifier_gaps: tuple[str, ...],
    changes: _Changes,
    part_diff: dict,
) -> dict:
    unapproved = changes.classification_counts["UNAPPROVED_LOSS"]
    fixture_gap_count = changes.classification_counts["FIXTURE_GAP"]
    verifier_gap_count = changes.classification_counts["VERIFIER_GAP"]
    blocking_issue_count = unapproved + fixture_gap_count + verifier_gap_count
    if fixture_gap_count:
        recommendation = "Repair or replace the fixture data before treating this comparison as product evidence."
    elif verifier_gap_count:
        recommendation = "Extend the verifier before treating this comparison as complete preservation evidence."
    elif unapproved:
        recommendation = "Do not replace the reference workbook until every unapproved difference is fixed or explicitly classified."
    elif changes.classification_counts["APPROVED_NORMALIZATION"]:
        recommendation = "Only approved normalizations and requested changes were detected."
    elif changes.classification_counts["REQUESTED"]:
        recommendation = "Only requested changes were detected."
    else:
        recommendation = "No semantic preservation differences were detected."
    return {
        "schema_version": 2,
        "before_path": before_path,
        "after_path": after_path,
        "before_sha256": before_sha256,
        "after_sha256": after_sha256,
        "before_size": before_size,
        "after_size": after_size,
        "equivalent": changes.total == 0,
        "preservation_ok": blocking_issue_count == 0,
        "recommendation": recommendation,
        "severity_counts": changes.severity_counts,
        "classification_counts": changes.classification_counts,
        "requested_paths": list(requested),
        "approved_normalizations": [rule["path"] for rule in approved_rules],
        "approved_normalization_rules": list(approved_rules),
        "normalization_evidence_complete": all(rule["complete"] for rule in approved_rules),
        "fixture_gap_paths": list(fixture_gaps),
        "verifier_gap_paths": list(verifier_gaps),
        "unapproved_difference_count": unapproved,
        "blocking_issue_count": blocking_issue_count,
        "part_diff": part_diff,
        "change_count": changes.total,
        "changes": changes.items,
        "truncated": changes.total > len(changes.items),
    }


def verify_xlsx_preservation(
    before: str,
    after: str,
    max_differences: int = 200,
    requested_paths: list[str] | tuple[str, ...] | None = None,
    approved_normalizations: list[str | dict] | tuple[str | dict, ...] | None = None,
    fixture_gap_paths: list[str] | tuple[str, ...] | None = None,
    verifier_gap_paths: list[str] | tuple[str, ...] | None = None,
    *,
    before_inspection: WorkbookInspection | None = None,
    after_inspection: WorkbookInspection | None = None,
) -> dict:
    """Compare two OOXML workbooks at package and semantic feature level."""
    if not 1 <= int(max_differences) <= 5000:
        raise ValueError("max_differences must be between 1 and 5000")
    requested = tuple(str(item) for item in (requested_paths or ()))
    approved_rules = _normalize_approved_normalization_rules(approved_normalizations)
    fixture_gaps = tuple(str(item) for item in (fixture_gap_paths or ()))
    verifier_gaps = tuple(str(item) for item in (verifier_gap_paths or ()))
    changes = _Changes(
        int(max_differences),
        requested,
        approved_rules,
        fixture_gaps,
        verifier_gaps,
    )
    before_path = str(Path(before).expanduser().resolve())
    after_path = str(Path(after).expanduser().resolve())
    if (before_inspection is None) != (after_inspection is None):
        raise ValueError(
            "before_inspection and after_inspection must be supplied together."
        )
    if before_inspection is None or after_inspection is None:
        left_inspection, right_inspection = inspect_workbook_pair(before_path, after_path)
    else:
        left_inspection, right_inspection = before_inspection, after_inspection
    if _path_key(left_inspection.path) != _path_key(before_path):
        raise ValueError("before_inspection does not match the before workbook path.")
    if _path_key(right_inspection.path) != _path_key(after_path):
        raise ValueError("after_inspection does not match the after workbook path.")
    for label, inspection in (
        ("before", left_inspection),
        ("after", right_inspection),
    ):
        if inspection.parts is None:
            errors = inspection.package_report.get("errors") or [
                f"Could not inspect the {label} workbook package."
            ]
            left_inspection.release_raw_parts()
            right_inspection.release_raw_parts()
            raise ValueError(str(errors[0]))
        if inspection.size is None or inspection.sha256 is None or inspection.state is None:
            left_inspection.release_raw_parts()
            right_inspection.release_raw_parts()
            raise ValueError(f"Missing fingerprint data for the {label} workbook.")

    before_size = left_inspection.size
    before_sha256 = left_inspection.sha256
    before_state = left_inspection.state
    after_size = right_inspection.size
    after_sha256 = right_inspection.sha256
    after_state = right_inspection.state
    if before_size == after_size and before_sha256 == after_sha256:
        try:
            _validate_package_xml(left_inspection.parts)
            part_diff = _byte_identical_part_diff(left_inspection.parts)
        finally:
            left_inspection.release_raw_parts()
            right_inspection.release_raw_parts()
        _assert_file_unchanged(before_path, before_state)
        if after_path != before_path:
            _assert_file_unchanged(after_path, after_state)
        return _verification_report(
            before_path,
            after_path,
            before_sha256,
            after_sha256,
            before_size,
            after_size,
            requested,
            approved_rules,
            fixture_gaps,
            verifier_gaps,
            changes,
            part_diff,
        )
    try:
        _validate_package_xml(left_inspection.parts)
        _validate_package_xml(right_inspection.parts)
        left = left_inspection.ensure_snapshot()
        right = right_inspection.ensure_snapshot()
        part_diff = _package_content_diff_from_inspections(
            left_inspection,
            right_inspection,
        )
    finally:
        left_inspection.release_raw_parts()
        right_inspection.release_raw_parts()

    left_wb = left["workbook"]
    right_wb = right["workbook"]
    for key, category, severity in (
        ("workbook_properties", "workbook_properties", "high"),
        ("workbook_protection", "workbook_protection", "critical"),
        ("calculation", "calculation", "high"),
        ("views", "workbook_views", "medium"),
        ("sheets", "sheet_inventory", "critical"),
        ("defined_names", "defined_names", "high"),
        ("external_references", "external_references", "critical"),
        ("extensions", "workbook_extensions", "high"),
    ):
        changes.add(f"workbook/{key}", category, severity, left_wb.get(key), right_wb.get(key))

    changes.add("styles/named_styles", "named_styles", "high", left["styles"]["named_styles"], right["styles"]["named_styles"])
    changes.add("styles/dxfs", "differential_styles", "high", left["styles"]["dxfs"], right["styles"]["dxfs"])
    changes.add("styles/table_styles", "table_styles", "high", left["styles"]["table_styles"], right["styles"]["table_styles"])
    changes.add("styles/colors", "indexed_colors", "high", left["styles"]["colors"], right["styles"]["colors"])
    changes.add("theme", "theme", "high", left["theme"], right["theme"])
    _compare_mapping(changes, "document_properties", "document_properties", "medium", left["document_properties"], right["document_properties"])

    left_sheets = left["worksheets"]
    right_sheets = right["worksheets"]
    for sheet_name in sorted(set(left_sheets) | set(right_sheets)):
        left_sheet = left_sheets.get(sheet_name)
        right_sheet = right_sheets.get(sheet_name)
        sheet_path = f"worksheets/{sheet_name}"
        if left_sheet is None or right_sheet is None:
            changes.add(sheet_path, "worksheet", "critical", left_sheet, right_sheet)
            continue
        for key, category, severity in (
            ("sheet_properties", "sheet_properties", "high"),
            ("sheet_views", "sheet_views", "medium"),
            ("row_defaults", "row_defaults", "high"),
            ("columns", "columns", "high"),
            ("merged_cells", "merged_cells", "high"),
            ("printing", "printing", "high"),
            ("filtering", "filtering", "high"),
            ("validation", "validation", "high"),
            ("protection", "sheet_protection", "critical"),
            ("ignored_errors", "ignored_errors", "medium"),
            ("hyperlinks", "hyperlinks", "high"),
            ("tables", "tables", "high"),
            ("comments", "comments", "high"),
            ("relationships", "relationships", "high"),
            ("extensions", "worksheet_extensions", "high"),
        ):
            changes.add(f"{sheet_path}/{key}", category, severity, left_sheet.get(key), right_sheet.get(key))
        _compare_mapping(changes, f"{sheet_path}/rows", "row_attributes", "high", left_sheet["rows"], right_sheet["rows"])
        _compare_cells(changes, sheet_name, left_sheet["cells"], right_sheet["cells"])

    left_advanced = left["advanced_parts"]
    right_advanced = right["advanced_parts"]
    for name in sorted(set(left_advanced) | set(right_advanced)):
        severity = "critical" if name.startswith(("xl/vbaProject", "_xmlsignatures/", "xl/activeX/", "xl/embeddings/")) else "high"
        changes.add(f"package/{name}", "advanced_part", severity, left_advanced.get(name), right_advanced.get(name))

    for key, value in part_diff["relationship_changes"]["added"].items():
        changes.add(f"package/relationships/{key}", "relationships", "high", None, value)
    for key, value in part_diff["relationship_changes"]["removed"].items():
        changes.add(f"package/relationships/{key}", "relationships", "high", value, None)
    for key, value in part_diff["relationship_changes"]["modified"].items():
        changes.add(f"package/relationships/{key}", "relationships", "high", value["before"], value["after"])
    for key, value in part_diff["content_type_changes"]["added"].items():
        changes.add(f"package/content_types/{key}", "content_types", "high", None, value)
    for key, value in part_diff["content_type_changes"]["removed"].items():
        changes.add(f"package/content_types/{key}", "content_types", "high", value, None)
    for key, value in part_diff["content_type_changes"]["modified"].items():
        changes.add(f"package/content_types/{key}", "content_types", "high", value["before"], value["after"])

    _assert_file_unchanged(before_path, before_state)
    if after_path != before_path:
        _assert_file_unchanged(after_path, after_state)
    return _verification_report(
        before_path,
        after_path,
        before_sha256,
        after_sha256,
        before_size,
        after_size,
        requested,
        approved_rules,
        fixture_gaps,
        verifier_gaps,
        changes,
        part_diff,
    )


def cleanup_excel_backups(now: datetime | None = None) -> dict:
    """Delete expired backup payloads and sidecars from the managed backup directory."""
    current = _utc_now(now)
    directory = _backup_directory()
    removed_backups = 0
    removed_sidecars = 0
    errors = []
    if not directory.exists():
        return {"directory": str(directory), "removed_backups": 0, "removed_sidecars": 0, "errors": []}
    for sidecar in directory.glob(f"{BACKUP_PREFIX}*.json"):
        try:
            record = json.loads(sidecar.read_text(encoding="utf-8"))
            expires_at = _parse_iso(record["expires_at"])
            if expires_at > current:
                continue
            backup_path = Path(record.get("backup_path", ""))
            if backup_path.is_file():
                backup_path.unlink()
                removed_backups += 1
            sidecar.unlink(missing_ok=True)
            removed_sidecars += 1
        except Exception as exc:
            errors.append(f"{sidecar}: {exc}")
    cutoff = current.timestamp() - BACKUP_RETENTION_DAYS * 24 * 60 * 60
    for backup in directory.glob(f"{BACKUP_PREFIX}*"):
        if backup.suffix == ".json" or not backup.is_file():
            continue
        sidecar = Path(str(backup) + ".json")
        try:
            if not sidecar.exists() and backup.stat().st_mtime < cutoff:
                backup.unlink()
                removed_backups += 1
        except OSError as exc:
            errors.append(f"{backup}: {exc}")
    return {
        "directory": str(directory),
        "removed_backups": removed_backups,
        "removed_sidecars": removed_sidecars,
        "errors": errors,
    }


def create_excel_backup(reference_path: str, saved_path: str) -> dict:
    """Copy the pre-save reference workbook and write a two-day retention sidecar."""
    cleanup_excel_backups()
    reference = Path(reference_path).expanduser().resolve()
    if not reference.is_file():
        raise FileNotFoundError(f"Cannot create pre-save backup; reference file not found: {reference}")
    saved = Path(saved_path).expanduser().resolve()
    directory = _backup_directory()
    directory.mkdir(parents=True, exist_ok=True)
    created = _utc_now()
    expires = created + timedelta(days=BACKUP_RETENTION_DAYS)
    token = uuid.uuid4().hex[:12]
    backup = directory / f"{BACKUP_PREFIX}{created.strftime('%Y%m%dT%H%M%SZ')}-{token}-{reference.name}"
    sidecar = Path(str(backup) + ".json")
    try:
        shutil.copy2(reference, backup)
        record = {
            "schema_version": 1,
            "backup_path": str(backup),
            "sidecar_path": str(sidecar),
            "reference_path": str(reference),
            "saved_path": str(saved),
            "created_at": _iso(created),
            "expires_at": _iso(expires),
            "retention_days": BACKUP_RETENTION_DAYS,
            "size": backup.stat().st_size,
            "sha256": _sha256(backup.read_bytes()),
        }
        temporary = Path(str(sidecar) + ".tmp")
        temporary.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporary, sidecar)
        return record
    except Exception:
        backup.unlink(missing_ok=True)
        sidecar.unlink(missing_ok=True)
        Path(str(sidecar) + ".tmp").unlink(missing_ok=True)
        raise


def discard_excel_backup(record: dict | None) -> None:
    """Remove a just-created backup when its corresponding save fails."""
    if not record:
        return
    for key in ("backup_path", "sidecar_path"):
        value = record.get(key)
        if value:
            Path(value).unlink(missing_ok=True)


def find_latest_excel_backup(saved_path: str) -> dict | None:
    """Return the newest unexpired backup record associated with a saved path."""
    cleanup_excel_backups()
    directory = _backup_directory()
    if not directory.exists():
        return None
    expected = _path_key(saved_path)
    candidates = []
    for sidecar in directory.glob(f"{BACKUP_PREFIX}*.json"):
        try:
            record = json.loads(sidecar.read_text(encoding="utf-8"))
            if _path_key(record.get("saved_path", "")) != expected:
                continue
            if _parse_iso(record["expires_at"]) <= _utc_now():
                continue
            if not Path(record["backup_path"]).is_file():
                continue
            record["sidecar_path"] = str(sidecar)
            candidates.append(record)
        except Exception:
            continue
    if not candidates:
        return None
    return max(candidates, key=lambda item: _parse_iso(item["created_at"]))


def start_excel_backup_cleanup() -> None:
    """Start one daemon that removes expired backups hourly while the server runs."""
    global _CLEANUP_STARTED
    with _CLEANUP_LOCK:
        if _CLEANUP_STARTED:
            return
        _CLEANUP_STARTED = True

    def worker() -> None:
        while True:
            try:
                cleanup_excel_backups()
            except Exception:
                pass
            threading.Event().wait(_CLEANUP_INTERVAL_SECONDS)

    threading.Thread(target=worker, name="excel-backup-cleanup", daemon=True).start()
