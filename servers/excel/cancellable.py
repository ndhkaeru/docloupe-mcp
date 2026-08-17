from __future__ import annotations

import base64
import hashlib
import importlib
import importlib.metadata
import json
import os
import pickle
import shutil
import subprocess
import sys
import sysconfig
import tempfile
import threading
import time
import traceback
from datetime import date, datetime, time as datetime_time
from pathlib import Path
from typing import Any

import anyio


_DEFAULT_TIMEOUT_SECONDS = 240.0
_MIN_TIMEOUT_SECONDS = 0.1
_MAX_TIMEOUT_SECONDS = 3600.0
_DEFAULT_MAX_WORKERS = 1
_DEFAULT_MAX_RESULT_BYTES = 64 * 1024 * 1024
_DEFAULT_MAX_LOAD_RESULT_BYTES = 512 * 1024 * 1024
_MIN_MAX_RESULT_BYTES = 1024 * 1024
_MAX_MAX_RESULT_BYTES = 2 * 1024 * 1024 * 1024
_DEFAULT_MAX_INPUT_BYTES = 512 * 1024 * 1024
_MIN_MAX_INPUT_BYTES = 1024 * 1024
_MAX_MAX_INPUT_BYTES = 2 * 1024 * 1024 * 1024
_MAX_PAYLOAD_BYTES = 1024 * 1024
_MAX_METADATA_BYTES = 16 * 1024
_MAX_REQUEST_BYTES = _MAX_PAYLOAD_BYTES + _MAX_METADATA_BYTES
_MAX_ERROR_LOG_BYTES = 16 * 1024
_POLL_INTERVAL_SECONDS = 0.05
_NORMAL_EXIT_GRACE_SECONDS = 3.0
_TERMINATE_GRACE_SECONDS = 2.0
_KILL_GRACE_SECONDS = 2.0
_WORKSPACE_REMOVE_DELAYS = (0.0, 0.05, 0.1, 0.25, 0.5, 1.0, 2.0)
_WORKER_FLAG = "--excel-worker"
_REQUEST_FILE = "request.json"
_INPUT_FILE = "input.pkl"
_RESULT_FILE = "result.json"
_STATUS_FILE = "status.json"
_ERROR_LOG_FILE = "worker.stderr.log"
_SUPPORTED_OPERATIONS = frozenset({
    "verify_preservation",
    "save_stage",
    "load",
    "_runtime_probe",
    "_test_sleep",
    "_test_large_result",
})
_MISSING = object()
_RESULT_TYPE_KEY = "__docloupe_result_type__"
_RESULT_VALUE_KEY = "value"


class HeavyOperationError(RuntimeError):
    def __init__(self, details: dict[str, Any]):
        self.details = details
        super().__init__(json.dumps(details, ensure_ascii=False, sort_keys=True))


class WorkerProcessError(HeavyOperationError):
    pass


class WorkerProcessTimeout(HeavyOperationError, TimeoutError):
    pass


class WorkerProcessCancelled(HeavyOperationError):
    pass


class _DeadlineExceeded(Exception):
    def __init__(self, phase: str):
        self.phase = phase
        super().__init__(phase)


class _WorkerReportedError(Exception):
    def __init__(self, error_type: str, message: str):
        self.error_type = error_type
        self.message = message
        super().__init__(message)


class _WorkerExited(Exception):
    def __init__(self, exit_code: int | None, message: str | None = None):
        self.exit_code = exit_code
        self.message = message
        super().__init__(str(exit_code))


class _ResultArtifactError(Exception):
    pass


def _configured_timeout_seconds(operation: str) -> float:
    operation_variable = {
        "verify_preservation": "EXCEL_MCP_VERIFY_TIMEOUT_SECONDS",
        "save_stage": "EXCEL_MCP_SAVE_TIMEOUT_SECONDS",
        "load": "EXCEL_MCP_LOAD_TIMEOUT_SECONDS",
    }.get(operation)
    raw = os.environ.get(operation_variable) if operation_variable else None
    if raw is None:
        raw = os.environ.get("EXCEL_MCP_HEAVY_TIMEOUT_SECONDS")
    if raw is None:
        return _DEFAULT_TIMEOUT_SECONDS
    try:
        value = float(raw)
    except ValueError:
        return _DEFAULT_TIMEOUT_SECONDS
    return min(max(value, _MIN_TIMEOUT_SECONDS), _MAX_TIMEOUT_SECONDS)


def _configured_max_workers() -> int:
    raw = os.environ.get("EXCEL_MCP_MAX_HEAVY_WORKERS")
    if raw is None:
        return _DEFAULT_MAX_WORKERS
    try:
        value = int(raw)
    except ValueError:
        return _DEFAULT_MAX_WORKERS
    return min(max(value, 1), 8)


def _configured_max_result_bytes(operation: str) -> int:
    if operation == "load":
        variable = "EXCEL_MCP_MAX_LOAD_RESULT_BYTES"
        default = _DEFAULT_MAX_LOAD_RESULT_BYTES
    else:
        variable = "EXCEL_MCP_MAX_RESULT_BYTES"
        default = _DEFAULT_MAX_RESULT_BYTES
    raw = os.environ.get(variable)
    if raw is None:
        return default
    try:
        value = int(raw)
    except ValueError:
        return default
    return min(max(value, _MIN_MAX_RESULT_BYTES), _MAX_MAX_RESULT_BYTES)


def _configured_max_input_bytes() -> int:
    raw = os.environ.get("EXCEL_MCP_MAX_INPUT_BYTES")
    if raw is None:
        return _DEFAULT_MAX_INPUT_BYTES
    try:
        value = int(raw)
    except ValueError:
        return _DEFAULT_MAX_INPUT_BYTES
    return min(max(value, _MIN_MAX_INPUT_BYTES), _MAX_MAX_INPUT_BYTES)


_HEAVY_SEMAPHORE = threading.BoundedSemaphore(_configured_max_workers())
_ACTIVE_WORKER_PIDS: set[int] = set()
_ACTIVE_WORKER_LOCK = threading.Lock()


def active_worker_pids() -> tuple[int, ...]:
    with _ACTIVE_WORKER_LOCK:
        return tuple(sorted(_ACTIVE_WORKER_PIDS))


def resolve_timeout_seconds(value: float | None, operation: str = "heavy_operation") -> float:
    timeout = _configured_timeout_seconds(operation) if value is None else float(value)
    if not _MIN_TIMEOUT_SECONDS <= timeout <= _MAX_TIMEOUT_SECONDS:
        raise ValueError(
            f"timeout_seconds must be between {_MIN_TIMEOUT_SECONDS} and "
            f"{_MAX_TIMEOUT_SECONDS}."
        )
    return timeout


def _sanitize_message(value: object, limit: int = 2000) -> str:
    message = " ".join(str(value).split())
    return message[:limit] if message else "Worker operation failed."


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _json_bytes(value: object) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        default=str,
        separators=(",", ":"),
    ).encode("utf-8")


def _result_json_default(value: object) -> dict[str, object]:
    if isinstance(value, datetime):
        return {
            _RESULT_TYPE_KEY: "datetime",
            _RESULT_VALUE_KEY: value.isoformat(),
            "fold": value.fold,
        }
    if isinstance(value, date):
        return {_RESULT_TYPE_KEY: "date", _RESULT_VALUE_KEY: value.isoformat()}
    if isinstance(value, datetime_time):
        return {
            _RESULT_TYPE_KEY: "time",
            _RESULT_VALUE_KEY: value.isoformat(),
            "fold": value.fold,
        }
    if isinstance(value, bytes):
        return {
            _RESULT_TYPE_KEY: "bytes",
            _RESULT_VALUE_KEY: base64.b64encode(value).decode("ascii"),
        }
    raise TypeError(f"Unsupported worker result type: {type(value).__name__}")


def _result_json_object_hook(value: dict[str, Any]) -> object:
    result_type = value.get(_RESULT_TYPE_KEY)
    if not isinstance(result_type, str) or _RESULT_VALUE_KEY not in value:
        return value
    encoded = value[_RESULT_VALUE_KEY]
    if not isinstance(encoded, str):
        raise ValueError("Tagged worker result value must be a string.")
    if result_type == "datetime":
        decoded = datetime.fromisoformat(encoded)
        return decoded.replace(fold=int(value.get("fold", 0)))
    if result_type == "date":
        return date.fromisoformat(encoded)
    if result_type == "time":
        decoded = datetime_time.fromisoformat(encoded)
        return decoded.replace(fold=int(value.get("fold", 0)))
    if result_type == "bytes":
        return base64.b64decode(encoded.encode("ascii"), validate=True)
    raise ValueError(f"Unsupported tagged worker result type: {result_type}")


def _result_json_bytes(value: object) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        default=_result_json_default,
        separators=(",", ":"),
    ).encode("utf-8")


def _result_json_loads(raw: bytes) -> Any:
    return json.loads(raw.decode("utf-8"), object_hook=_result_json_object_hook)


def _validated_payload(payload: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise TypeError("operation payload must be a dictionary")
    encoded = _json_bytes(payload)
    if len(encoded) > _MAX_PAYLOAD_BYTES:
        raise ValueError(f"operation payload exceeds {_MAX_PAYLOAD_BYTES} bytes")
    return json.loads(encoded.decode("utf-8"))


def _excel_main_module():
    for module_name in ("main", "__main__"):
        module = sys.modules.get(module_name)
        if module is not None and hasattr(module, "excel_verify_preservation"):
            return module
    return importlib.import_module("main")


def _execute_named_operation(
    operation: str,
    payload: dict[str, Any],
    input_data: Any = _MISSING,
) -> Any:
    if operation == "verify_preservation":
        return _excel_main_module().excel_verify_preservation(**payload)
    if operation == "save_stage":
        if input_data is _MISSING:
            raise ValueError("save_stage requires a validated input artifact")
        return _excel_main_module()._save_stage_operation(input_data, **payload)
    if operation == "load":
        return _excel_main_module()._load_worker_operation(**payload)
    if operation == "_runtime_probe":
        return {
            "pid": os.getpid(),
            "executable": sys.executable,
            "prefix": sys.prefix,
            "base_prefix": sys.base_prefix,
            "mcp": importlib.metadata.version("mcp"),
            "openpyxl": importlib.metadata.version("openpyxl"),
            "sys_path": list(sys.path),
        }
    if operation == "_test_sleep":
        seconds = float(payload.get("seconds", 0.0))
        if not 0.0 <= seconds <= 60.0:
            raise ValueError("seconds must be between 0 and 60")
        time.sleep(seconds)
        return {"slept_seconds": seconds, "pid": os.getpid()}
    if operation == "_test_large_result":
        size = int(payload.get("size", 0))
        if not 0 <= size <= 128 * 1024 * 1024:
            raise ValueError("size must be between 0 and 134217728")
        return {"value": "x" * size, "pid": os.getpid()}
    raise ValueError(f"Unsupported heavy operation: {operation}")


def _atomic_write(path: Path, data: bytes) -> None:
    temporary = path.with_name(f"{path.name}.tmp")
    temporary.write_bytes(data)
    os.replace(temporary, path)


def _write_status(path: Path, metadata: dict[str, Any]) -> None:
    encoded = _json_bytes(metadata)
    if len(encoded) > _MAX_METADATA_BYTES:
        encoded = _json_bytes({
            "status": "error",
            "error_type": "MetadataTooLarge",
            "message": "Worker metadata exceeded the status-file limit.",
        })
    _atomic_write(path, encoded)


def _write_input_artifact(
    workspace: Path,
    input_data: Any,
    max_input_bytes: int,
) -> dict[str, Any] | None:
    if input_data is _MISSING:
        return None
    encoded = pickle.dumps(input_data, protocol=pickle.HIGHEST_PROTOCOL)
    if len(encoded) > max_input_bytes:
        raise ValueError(
            f"Worker input exceeds the configured limit of {max_input_bytes} bytes."
        )
    _atomic_write(workspace / _INPUT_FILE, encoded)
    return {
        "format": "pickle-v1",
        "size": len(encoded),
        "sha256": _sha256(encoded),
    }


def _read_input_artifact(
    workspace: Path,
    metadata: object,
    max_input_bytes: int,
) -> Any:
    if metadata is None:
        return _MISSING
    if not isinstance(metadata, dict) or metadata.get("format") != "pickle-v1":
        raise ValueError("Worker input metadata is invalid.")
    artifact = workspace / _INPUT_FILE
    if not artifact.is_file():
        raise ValueError("Worker input artifact is missing.")
    raw = artifact.read_bytes()
    expected_size = metadata.get("size")
    expected_sha256 = metadata.get("sha256")
    if len(raw) > max_input_bytes:
        raise ValueError("Worker input artifact exceeds the configured limit.")
    if not isinstance(expected_size, int) or expected_size != len(raw):
        raise ValueError("Worker input artifact size does not match metadata.")
    if not isinstance(expected_sha256, str) or expected_sha256 != _sha256(raw):
        raise ValueError("Worker input artifact hash does not match metadata.")
    try:
        return pickle.loads(raw)
    except Exception as exc:
        raise ValueError("Worker input artifact could not be decoded.") from exc


def _write_worker_request(
    workspace: Path,
    operation: str,
    payload: dict[str, Any],
    max_result_bytes: int,
    max_input_bytes: int,
    input_metadata: dict[str, Any] | None,
) -> None:
    encoded = _json_bytes({
        "version": 1,
        "operation": operation,
        "payload": payload,
        "max_result_bytes": max_result_bytes,
        "max_input_bytes": max_input_bytes,
        "input": input_metadata,
    })
    if len(encoded) > _MAX_REQUEST_BYTES:
        raise ValueError(f"worker request exceeds {_MAX_REQUEST_BYTES} bytes")
    _atomic_write(workspace / _REQUEST_FILE, encoded)


def _read_worker_request(workspace: Path) -> tuple[str, dict[str, Any], int, Any]:
    request_path = workspace / _REQUEST_FILE
    raw = request_path.read_bytes()
    if len(raw) > _MAX_REQUEST_BYTES:
        raise ValueError("Worker request exceeds the configured limit.")
    try:
        request = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("Worker request is not valid UTF-8 JSON.") from exc
    if not isinstance(request, dict) or request.get("version") != 1:
        raise ValueError("Worker request version is unsupported.")
    operation = request.get("operation")
    if operation not in _SUPPORTED_OPERATIONS:
        raise ValueError(f"Unsupported heavy operation: {operation}")
    payload = _validated_payload(request.get("payload"))
    max_result_bytes = request.get("max_result_bytes")
    if (
        not isinstance(max_result_bytes, int)
        or isinstance(max_result_bytes, bool)
        or not _MIN_MAX_RESULT_BYTES <= max_result_bytes <= _MAX_MAX_RESULT_BYTES
    ):
        raise ValueError("Worker result limit is invalid.")
    max_input_bytes = request.get("max_input_bytes", _DEFAULT_MAX_INPUT_BYTES)
    if (
        not isinstance(max_input_bytes, int)
        or isinstance(max_input_bytes, bool)
        or not _MIN_MAX_INPUT_BYTES <= max_input_bytes <= _MAX_MAX_INPUT_BYTES
    ):
        raise ValueError("Worker input limit is invalid.")
    input_data = _read_input_artifact(
        workspace,
        request.get("input"),
        max_input_bytes,
    )
    return operation, payload, max_result_bytes, input_data


def run_worker_cli(workspace_path: str) -> int:
    workspace = Path(workspace_path).resolve(strict=True)
    if not workspace.is_dir():
        raise NotADirectoryError(str(workspace))
    artifact = workspace / _RESULT_FILE
    status_path = workspace / _STATUS_FILE
    temporary_artifact = artifact.with_name(f"{artifact.name}.tmp")
    temporary_status = status_path.with_name(f"{status_path.name}.tmp")
    for path in (artifact, status_path, temporary_artifact, temporary_status):
        path.unlink(missing_ok=True)
    try:
        operation, payload, max_result_bytes, input_data = _read_worker_request(workspace)
        operation_started_at = time.perf_counter()
        result = _execute_named_operation(operation, payload, input_data)
        operation_seconds = time.perf_counter() - operation_started_at
        encode_started_at = time.perf_counter()
        encoded = _result_json_bytes(result)
        encode_seconds = time.perf_counter() - encode_started_at
        if len(encoded) > max_result_bytes:
            raise ValueError(
                f"Worker result exceeds the configured limit of {max_result_bytes} bytes."
            )
        write_started_at = time.perf_counter()
        _atomic_write(artifact, encoded)
        write_seconds = time.perf_counter() - write_started_at
        _write_status(status_path, {
            "status": "ok",
            "size": len(encoded),
            "sha256": _sha256(encoded),
            "metrics": {
                "operation_seconds": round(operation_seconds, 6),
                "json_encode_seconds": round(encode_seconds, 6),
                "artifact_write_seconds": round(write_seconds, 6),
                "artifact_bytes": len(encoded),
            },
        })
        return 0
    except BaseException as exc:
        traceback.print_exc(file=sys.stderr)
        temporary_artifact.unlink(missing_ok=True)
        artifact.unlink(missing_ok=True)
        try:
            _write_status(status_path, {
                "status": "error",
                "error_type": type(exc).__name__,
                "message": _sanitize_message(exc),
            })
        except BaseException:
            traceback.print_exc(file=sys.stderr)
            return 2
        return 1


def _worker_command(workspace: Path) -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, _WORKER_FLAG, str(workspace)]
    entrypoint = Path(__file__).resolve().with_name("main.py")
    if os.name == "nt" and sys.prefix != sys.base_prefix:
        base_interpreter = Path(sys.base_prefix) / Path(sys.executable).name
        site_packages = Path(sysconfig.get_paths()["purelib"])
        if base_interpreter.is_file() and site_packages.is_dir():
            bootstrap = (
                "import os,runpy,site,sys;"
                "site_path=os.path.abspath(sys.argv.pop(1));"
                "site.addsitedir(site_path);"
                "target=os.path.normcase(site_path);"
                "cleaned=[];"
                "exec(\"for item in sys.path:\\n"
                "    if not item:\\n"
                "        cleaned.append(item); continue\\n"
                "    absolute=os.path.abspath(item)\\n"
                "    normalized=os.path.normcase(absolute)\\n"
                "    lowered=normalized.lower()\\n"
                "    is_site=('site-packages' in lowered or 'dist-packages' in lowered)\\n"
                "    try:\\n"
                "        in_target=os.path.commonpath([target, normalized]) == target\\n"
                "    except ValueError:\\n"
                "        in_target=False\\n"
                "    if is_site and not in_target:\\n"
                "        continue\\n"
                "    if normalized != target:\\n"
                "        cleaned.append(item)\");"
                "sys.path[:]=[site_path,*cleaned];"
                "entrypoint=sys.argv.pop(1);"
                "sys.argv[0]=entrypoint;"
                "runpy.run_path(entrypoint,run_name='__main__')"
            )
            return [
                str(base_interpreter),
                "-E",
                "-s",
                "-c",
                bootstrap,
                str(site_packages),
                str(entrypoint),
                _WORKER_FLAG,
                str(workspace),
            ]
    return [sys.executable, str(entrypoint), _WORKER_FLAG, str(workspace)]


def _start_worker(
    workspace: Path,
    operation: str,
    payload: dict[str, Any],
    max_result_bytes: int,
    max_input_bytes: int,
    input_data: Any,
) -> subprocess.Popen:
    input_metadata = _write_input_artifact(workspace, input_data, max_input_bytes)
    _write_worker_request(
        workspace,
        operation,
        payload,
        max_result_bytes,
        max_input_bytes,
        input_metadata,
    )
    error_log = workspace / _ERROR_LOG_FILE
    creationflags = 0
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    environment = os.environ.copy()
    environment["EXCEL_MCP_WORKER_PROCESS"] = "1"
    if getattr(sys, "frozen", False):
        environment["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    with error_log.open("wb") as error_stream:
        return subprocess.Popen(
            _worker_command(workspace),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=error_stream,
            env=environment,
            close_fds=True,
            creationflags=creationflags,
        )


def _terminate_process_tree(process: subprocess.Popen) -> None:
    if process.pid is None or process.poll() is not None:
        return
    if os.name == "nt":
        system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
        taskkill = system_root / "System32" / "taskkill.exe"
        command = str(taskkill) if taskkill.is_file() else "taskkill.exe"
        completed = subprocess.run(
            [command, "/PID", str(process.pid), "/T", "/F"],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if completed.returncode == 0 or process.poll() is not None:
            return
    process.terminate()


def _finish_process(process: subprocess.Popen, terminate_first: bool) -> bool:
    if process.pid is None:
        return True
    if terminate_first:
        _terminate_process_tree(process)
    try:
        process.wait(
            timeout=_TERMINATE_GRACE_SECONDS if terminate_first else _NORMAL_EXIT_GRACE_SECONDS
        )
    except subprocess.TimeoutExpired:
        _terminate_process_tree(process)
        try:
            process.wait(timeout=_TERMINATE_GRACE_SECONDS)
        except subprocess.TimeoutExpired:
            if process.poll() is None:
                process.kill()
            try:
                process.wait(timeout=_KILL_GRACE_SECONDS)
            except subprocess.TimeoutExpired:
                pass
    return process.poll() is not None


def _remove_workspace(path: Path | None) -> bool:
    if path is None or not path.exists():
        return True
    for attempt, delay in enumerate(_WORKSPACE_REMOVE_DELAYS):
        if delay:
            time.sleep(delay)
        try:
            shutil.rmtree(path)
        except FileNotFoundError:
            return True
        except OSError:
            if attempt == len(_WORKSPACE_REMOVE_DELAYS) - 1:
                return False
            continue
        return not path.exists()
    return not path.exists()


def _remove_failure_paths(paths: tuple[Path, ...]) -> bool:
    all_removed = True
    for path in paths:
        if not path.exists():
            continue
        removed = False
        for attempt, delay in enumerate(_WORKSPACE_REMOVE_DELAYS):
            if delay:
                time.sleep(delay)
            try:
                if path.is_dir() and not path.is_symlink():
                    shutil.rmtree(path)
                else:
                    path.unlink(missing_ok=True)
            except OSError:
                if attempt == len(_WORKSPACE_REMOVE_DELAYS) - 1:
                    break
                continue
            removed = not path.exists()
            if removed:
                break
        all_removed = all_removed and removed
    return all_removed


def _read_status(path: Path) -> dict[str, Any]:
    raw = path.read_bytes()
    if len(raw) > _MAX_METADATA_BYTES:
        raise _ResultArtifactError("Worker status exceeds the configured limit.")
    try:
        metadata = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise _ResultArtifactError("Worker status is not valid UTF-8 JSON.") from exc
    if not isinstance(metadata, dict):
        raise _ResultArtifactError("Worker status must be a JSON object.")
    return metadata


def _read_error_log(workspace: Path | None) -> str | None:
    if workspace is None:
        return None
    path = workspace / _ERROR_LOG_FILE
    if not path.is_file():
        return None
    try:
        size = path.stat().st_size
        with path.open("rb") as stream:
            if size > _MAX_ERROR_LOG_BYTES:
                stream.seek(-_MAX_ERROR_LOG_BYTES, os.SEEK_END)
            raw = stream.read(_MAX_ERROR_LOG_BYTES)
        return _sanitize_message(raw.decode("utf-8", errors="replace"))
    except OSError:
        return None


def _read_result_artifact(
    artifact: Path,
    metadata: dict[str, Any],
    max_result_bytes: int,
) -> tuple[Any, dict[str, float | int]]:
    read_started_at = time.perf_counter()
    if not artifact.is_file():
        raise _ResultArtifactError("Worker result artifact is missing.")
    raw = artifact.read_bytes()
    read_seconds = time.perf_counter() - read_started_at
    validate_started_at = time.perf_counter()
    expected_size = metadata.get("size")
    expected_sha256 = metadata.get("sha256")
    if len(raw) > max_result_bytes:
        raise _ResultArtifactError("Worker result artifact exceeds the configured limit.")
    if not isinstance(expected_size, int) or expected_size != len(raw):
        raise _ResultArtifactError("Worker result artifact size does not match metadata.")
    if not isinstance(expected_sha256, str) or expected_sha256 != _sha256(raw):
        raise _ResultArtifactError("Worker result artifact hash does not match metadata.")
    validate_seconds = time.perf_counter() - validate_started_at
    decode_started_at = time.perf_counter()
    try:
        result = _result_json_loads(raw)
    except (UnicodeDecodeError, json.JSONDecodeError, TypeError, ValueError) as exc:
        raise _ResultArtifactError("Worker result artifact is not valid tagged UTF-8 JSON.") from exc
    decode_seconds = time.perf_counter() - decode_started_at
    return result, {
        "artifact_read_seconds": round(read_seconds, 6),
        "artifact_validate_seconds": round(validate_seconds, 6),
        "json_decode_seconds": round(decode_seconds, 6),
        "artifact_bytes": len(raw),
    }


async def _acquire_heavy_slot(deadline: float) -> None:
    while not _HEAVY_SEMAPHORE.acquire(blocking=False):
        if anyio.current_time() >= deadline:
            raise _DeadlineExceeded("waiting_for_slot")
        await anyio.sleep(_POLL_INTERVAL_SECONDS)


def _failure_details(
    code: str,
    operation: str,
    timeout_seconds: float,
    elapsed_seconds: float,
    worker_started: bool,
    worker_stopped: bool,
    staging_removed: bool,
    **extra: Any,
) -> dict[str, Any]:
    return {
        "code": code,
        "operation": operation,
        "timeout_seconds": timeout_seconds,
        "elapsed_seconds": round(elapsed_seconds, 3),
        "worker_started": worker_started,
        "worker_stopped": worker_stopped,
        "staging_removed": staging_removed,
        **extra,
    }


async def run_named_operation(
    operation: str,
    payload: dict[str, Any],
    *,
    timeout_seconds: float | None = None,
    operation_label: str | None = None,
    input_data: Any = _MISSING,
    failure_cleanup_paths: tuple[str | Path, ...] = (),
    return_metadata: bool = False,
) -> Any:
    if operation not in _SUPPORTED_OPERATIONS:
        raise ValueError(f"Unsupported heavy operation: {operation}")
    validated_payload = _validated_payload(payload)
    timeout = resolve_timeout_seconds(timeout_seconds, operation)
    max_result_bytes = _configured_max_result_bytes(operation)
    max_input_bytes = _configured_max_input_bytes()
    cleanup_paths = tuple(Path(path).expanduser().resolve() for path in failure_cleanup_paths)
    label = operation_label or operation
    started_at = time.perf_counter()
    deadline = anyio.current_time() + timeout
    acquired = False
    workspace: Path | None = None
    artifact: Path | None = None
    status_path: Path | None = None
    process: subprocess.Popen | None = None
    process_pid: int | None = None
    result: Any = _MISSING
    result_metrics: dict[str, float | int] = {}
    worker_metadata: dict[str, Any] = {}
    failure: BaseException | None = None
    worker_stopped = True
    workspace_removed = True
    failure_paths_removed = True
    staging_removed = True
    cleanup_error: str | None = None

    try:
        await _acquire_heavy_slot(deadline)
        acquired = True
        workspace = Path(tempfile.mkdtemp(prefix="docloupe-excel-worker-"))
        artifact = workspace / _RESULT_FILE
        status_path = workspace / _STATUS_FILE
        process = await anyio.to_thread.run_sync(
            _start_worker,
            workspace,
            operation,
            validated_payload,
            max_result_bytes,
            max_input_bytes,
            input_data,
        )
        process_pid = process.pid
        if process_pid is not None:
            with _ACTIVE_WORKER_LOCK:
                _ACTIVE_WORKER_PIDS.add(process_pid)

        while True:
            if status_path.is_file():
                metadata = _read_status(status_path)
                if metadata.get("status") == "ok":
                    worker_metadata = metadata
                    result, result_metrics = await anyio.to_thread.run_sync(
                        _read_result_artifact,
                        artifact,
                        metadata,
                        max_result_bytes,
                    )
                    break
                raise _WorkerReportedError(
                    str(metadata.get("error_type") or "WorkerError"),
                    _sanitize_message(metadata.get("message")),
                )
            exit_code = process.poll()
            if exit_code is not None:
                if status_path.is_file():
                    continue
                raise _WorkerExited(exit_code, _read_error_log(workspace))
            if anyio.current_time() >= deadline:
                raise _DeadlineExceeded("running")
            await anyio.sleep(_POLL_INTERVAL_SECONDS)
    except BaseException as exc:
        failure = exc
    finally:
        cancelled_type = anyio.get_cancelled_exc_class()
        terminate_first = failure is not None
        with anyio.CancelScope(shield=True):
            if process is not None:
                try:
                    worker_stopped = await anyio.to_thread.run_sync(
                        _finish_process,
                        process,
                        terminate_first,
                    )
                except BaseException as exc:
                    worker_stopped = False
                    cleanup_error = _sanitize_message(exc)
                if process_pid is not None:
                    with _ACTIVE_WORKER_LOCK:
                        _ACTIVE_WORKER_PIDS.discard(process_pid)
            workspace_removed = await anyio.to_thread.run_sync(_remove_workspace, workspace)
            if failure is not None and cleanup_paths:
                failure_paths_removed = await anyio.to_thread.run_sync(
                    _remove_failure_paths,
                    cleanup_paths,
                )
            staging_removed = workspace_removed and failure_paths_removed
            if acquired:
                _HEAVY_SEMAPHORE.release()

        elapsed = time.perf_counter() - started_at
        worker_started = process_pid is not None
        common = {
            "operation": label,
            "timeout_seconds": timeout,
            "elapsed_seconds": elapsed,
            "worker_started": worker_started,
            "worker_pid": process_pid,
            "worker_stopped": worker_stopped,
            "workspace_removed": workspace_removed,
            "failure_paths_removed": failure_paths_removed,
            "staging_removed": staging_removed,
        }
        if cleanup_error or not worker_stopped or not staging_removed:
            failure = WorkerProcessError(_failure_details(
                "HEAVY_OPERATION_CLEANUP_FAILED",
                **common,
                message=cleanup_error or "Worker or staging cleanup did not complete.",
            ))
        elif isinstance(failure, cancelled_type):
            failure = WorkerProcessCancelled(_failure_details(
                "HEAVY_OPERATION_CANCELLED",
                **common,
            ))
        elif isinstance(failure, _DeadlineExceeded):
            failure = WorkerProcessTimeout(_failure_details(
                "HEAVY_OPERATION_TIMEOUT",
                **common,
                phase=failure.phase,
            ))
        elif isinstance(failure, _WorkerReportedError):
            failure = WorkerProcessError(_failure_details(
                "HEAVY_OPERATION_FAILED",
                **common,
                error_type=failure.error_type,
                message=failure.message,
            ))
        elif isinstance(failure, _WorkerExited):
            failure = WorkerProcessError(_failure_details(
                "HEAVY_OPERATION_WORKER_EXIT",
                **common,
                exit_code=failure.exit_code,
                **({"message": failure.message} if failure.message else {}),
            ))
        elif isinstance(failure, _ResultArtifactError):
            failure = WorkerProcessError(_failure_details(
                "HEAVY_OPERATION_RESULT_INVALID",
                **common,
                message=str(failure),
            ))
        elif failure is not None and not isinstance(failure, HeavyOperationError):
            failure = WorkerProcessError(_failure_details(
                "HEAVY_OPERATION_SUPERVISOR_FAILED",
                **common,
                error_type=type(failure).__name__,
                message=_sanitize_message(failure),
            ))

    if failure is not None:
        raise failure
    if result is _MISSING:
        raise WorkerProcessError(_failure_details(
            "HEAVY_OPERATION_RESULT_MISSING",
            label,
            timeout,
            time.perf_counter() - started_at,
            process_pid is not None,
            worker_stopped,
            staging_removed,
        ))
    if return_metadata:
        worker_metrics = worker_metadata.get("metrics")
        if not isinstance(worker_metrics, dict):
            worker_metrics = {}
        return {
            "result": result,
            "metrics": {
                "worker": worker_metrics,
                "parent": result_metrics,
                "supervisor_seconds": round(time.perf_counter() - started_at, 6),
            },
        }
    return result
