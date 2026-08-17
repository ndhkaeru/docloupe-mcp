from __future__ import annotations

import ctypes
import importlib.metadata
import json
import os
import shutil
import sys
import sysconfig
from pathlib import Path
from types import SimpleNamespace

import anyio
import openpyxl
import pytest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers" / "excel"))

import cancellable  # noqa: E402
import main as M  # noqa: E402
from cancellable import (  # noqa: E402
    WorkerProcessCancelled,
    WorkerProcessTimeout,
    active_worker_pids,
    run_named_operation,
)


def _capture_worker_workspaces(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> list[Path]:
    created: list[Path] = []

    def fake_mkdtemp(suffix: str = "", prefix: str = "tmp", dir=None) -> str:
        del dir
        workspace = tmp_path / f"{prefix}{len(created)}{suffix}"
        workspace.mkdir(parents=True)
        created.append(workspace)
        return str(workspace)

    monkeypatch.setattr(
        cancellable,
        "tempfile",
        SimpleNamespace(mkdtemp=fake_mkdtemp),
    )
    return created


def _record_finished_processes(
    monkeypatch: pytest.MonkeyPatch,
) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    original = cancellable._finish_process

    def recording_finish(process, terminate_first: bool) -> bool:
        pid = process.pid
        stopped = original(process, terminate_first)
        records.append({
            "pid": pid,
            "terminate_first": terminate_first,
            "stopped": stopped,
        })
        return stopped

    monkeypatch.setattr(cancellable, "_finish_process", recording_finish)
    return records


def _process_is_running(pid: int) -> bool:
    if os.name != "nt":
        try:
            os.kill(pid, 0)
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        return True

    from ctypes import wintypes

    process_query_limited_information = 0x1000
    access_denied = 5
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = (
        wintypes.DWORD,
        wintypes.BOOL,
        wintypes.DWORD,
    )
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel32.CloseHandle.restype = wintypes.BOOL
    kernel32.GetExitCodeProcess.argtypes = (
        wintypes.HANDLE,
        ctypes.POINTER(wintypes.DWORD),
    )
    kernel32.GetExitCodeProcess.restype = wintypes.BOOL

    handle = kernel32.OpenProcess(process_query_limited_information, False, pid)
    if handle:
        exit_code = wintypes.DWORD()
        queried = kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code))
        kernel32.CloseHandle(handle)
        return bool(queried) and exit_code.value == 259
    return ctypes.get_last_error() == access_denied


async def _invoke_named_operation(
    operation: str,
    payload: dict,
    timeout_seconds: float,
    operation_label: str,
):
    return await run_named_operation(
        operation,
        payload,
        timeout_seconds=timeout_seconds,
        operation_label=operation_label,
    )


async def _wait_for_worker_pid(timeout_seconds: float = 10.0) -> int:
    with anyio.fail_after(timeout_seconds):
        while True:
            pids = active_worker_pids()
            if pids:
                return pids[0]
            await anyio.sleep(0.01)


async def _cancel_running_sleep() -> tuple[int, WorkerProcessCancelled]:
    outcome: dict[str, object] = {}

    async def invoke() -> None:
        try:
            await run_named_operation(
                "_test_sleep",
                {"seconds": 30.0},
                timeout_seconds=20.0,
                operation_label="host-cancel-regression",
            )
        except WorkerProcessCancelled as exc:
            outcome["error"] = exc

    async with anyio.create_task_group() as task_group:
        task_group.start_soon(invoke)
        outcome["pid"] = await _wait_for_worker_pid()
        task_group.cancel_scope.cancel()

    assert isinstance(outcome.get("pid"), int)
    assert isinstance(outcome.get("error"), WorkerProcessCancelled)
    return outcome["pid"], outcome["error"]


async def _verify_simple_workbook(before_path: Path, after_path: Path) -> str:
    return await M._excel_verify_preservation_tool(
        after_path=str(after_path),
        before_path=str(before_path),
        fixture_id="cancellation-wrapper",
        timeout_seconds=30.0,
    )


def _write_simple_workbook(path: Path) -> None:
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = "preserve me"
    workbook.save(path)
    workbook.close()


def test_run_named_operation_success_cleans_result_artifact(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspaces = _capture_worker_workspaces(monkeypatch, tmp_path)
    finished = _record_finished_processes(monkeypatch)

    result = anyio.run(
        _invoke_named_operation,
        "_test_sleep",
        {"seconds": 0.0},
        5.0,
        "success-regression",
    )

    assert result["slept_seconds"] == 0.0
    assert result["pid"] != os.getpid()
    assert finished == [{
        "pid": result["pid"],
        "terminate_first": False,
        "stopped": True,
    }]
    assert active_worker_pids() == ()
    assert not _process_is_running(result["pid"])
    assert len(workspaces) == 1
    assert not workspaces[0].exists()


def test_source_worker_uses_exact_virtualenv_dependencies() -> None:
    result = anyio.run(
        _invoke_named_operation,
        "_runtime_probe",
        {},
        10.0,
        "runtime-probe",
    )

    expected_site = Path(sysconfig.get_paths()["purelib"]).resolve()
    foreign_site_paths = []
    for raw_path in result["sys_path"]:
        if not raw_path:
            continue
        path = Path(raw_path).resolve()
        lowered = str(path).lower()
        if "site-packages" not in lowered and "dist-packages" not in lowered:
            continue
        try:
            path.relative_to(expected_site)
        except ValueError:
            foreign_site_paths.append(str(path))

    assert result["mcp"] == importlib.metadata.version("mcp") == "1.28.1"
    assert result["openpyxl"] == importlib.metadata.version("openpyxl") == "3.1.5"
    assert expected_site in {Path(item).resolve() for item in result["sys_path"] if item}
    assert foreign_site_paths == []


def test_deferred_mcp_cancellation_sends_one_structured_response() -> None:
    from mcp.shared.session import RequestResponder
    from mcp.types import (
        CallToolRequest,
        CallToolRequestParams,
        ClientRequest,
        ErrorData,
    )

    class FakeSession:
        def __init__(self) -> None:
            self.responses = []

        async def _send_response(self, request_id, response) -> None:
            self.responses.append((request_id, response))

    async def scenario() -> list:
        session = FakeSession()
        completed = []
        responder = RequestResponder(
            request_id=7,
            request_meta=None,
            request=ClientRequest(root=CallToolRequest(
                params=CallToolRequestParams(
                    name="excel_verify_preservation",
                    arguments={},
                )
            )),
            session=session,
            on_complete=lambda value: completed.append(value.request_id),
        )
        with responder:
            await responder.cancel()
            assert responder.cancelled is True
            assert responder._completed is False
            await responder.respond(ErrorData(
                code=0,
                message="structured cancellation",
                data={"worker_stopped": True, "staging_removed": True},
            ))
        assert completed == [7]
        return session.responses

    responses = anyio.run(scenario)
    assert len(responses) == 1
    assert responses[0][0] == 7
    assert responses[0][1].data == {
        "worker_stopped": True,
        "staging_removed": True,
    }


def test_timeout_stops_worker_and_returns_structured_details(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspaces = _capture_worker_workspaces(monkeypatch, tmp_path)
    finished = _record_finished_processes(monkeypatch)

    with pytest.raises(WorkerProcessTimeout) as caught:
        anyio.run(
            _invoke_named_operation,
            "_test_sleep",
            {"seconds": 30.0},
            0.2,
            "timeout-regression",
        )

    details = caught.value.details
    assert details["code"] == "HEAVY_OPERATION_TIMEOUT"
    assert details["operation"] == "timeout-regression"
    assert details["timeout_seconds"] == 0.2
    assert details["phase"] == "running"
    assert details["worker_started"] is True
    assert details["worker_stopped"] is True
    assert details["staging_removed"] is True
    assert details["elapsed_seconds"] >= 0.0
    assert len(finished) == 1
    assert finished[0]["terminate_first"] is True
    assert finished[0]["stopped"] is True
    assert isinstance(finished[0]["pid"], int)
    assert active_worker_pids() == ()
    assert not _process_is_running(finished[0]["pid"])
    assert len(workspaces) == 1
    assert not workspaces[0].exists()


def test_host_task_cancellation_stops_worker(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspaces = _capture_worker_workspaces(monkeypatch, tmp_path)
    finished = _record_finished_processes(monkeypatch)

    pid, error = anyio.run(_cancel_running_sleep)

    assert error.details["code"] == "HEAVY_OPERATION_CANCELLED"
    assert error.details["operation"] == "host-cancel-regression"
    assert error.details["timeout_seconds"] == 20.0
    assert error.details["worker_started"] is True
    assert error.details["worker_stopped"] is True
    assert error.details["staging_removed"] is True
    assert finished == [{
        "pid": pid,
        "terminate_first": True,
        "stopped": True,
    }]
    assert active_worker_pids() == ()
    assert not _process_is_running(pid)
    assert len(workspaces) == 1
    assert not workspaces[0].exists()


def test_timeout_releases_slot_for_next_fast_request(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspaces = _capture_worker_workspaces(monkeypatch, tmp_path)

    async def scenario() -> dict:
        with pytest.raises(WorkerProcessTimeout):
            await run_named_operation(
                "_test_sleep",
                {"seconds": 30.0},
                timeout_seconds=0.2,
                operation_label="slot-timeout-regression",
            )
        return await run_named_operation(
            "_test_sleep",
            {"seconds": 0.0},
            timeout_seconds=5.0,
            operation_label="slot-success-regression",
        )

    result = anyio.run(scenario)

    assert result["slept_seconds"] == 0.0
    assert active_worker_pids() == ()
    assert len(workspaces) == 2
    assert all(not workspace.exists() for workspace in workspaces)


def test_async_mcp_verify_wrapper_returns_valid_json(tmp_path: Path) -> None:
    before_path = tmp_path / "before.xlsx"
    after_path = tmp_path / "after.xlsx"
    _write_simple_workbook(before_path)
    shutil.copy2(before_path, after_path)

    raw_report = anyio.run(_verify_simple_workbook, before_path, after_path)
    report = json.loads(raw_report)

    assert report["fixture_id"] == "cancellation-wrapper"
    assert report["fixture_id_source"] == "provided"
    assert report["equivalent"] is True
    assert report["preservation_ok"] is True
    assert report["change_count"] == 0
    assert report["backup"] is None
    assert report["files"]["before"]["path"] == str(before_path.resolve())
    assert report["files"]["after"]["path"] == str(after_path.resolve())
    assert report["files"]["before"]["package_valid"] is True
    assert report["files"]["after"]["package_valid"] is True
    assert report["files"]["before"]["loadable"] is True
    assert report["files"]["after"]["loadable"] is True
    assert active_worker_pids() == ()


def test_workspace_cleanup_retries_transient_file_lock(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspace = tmp_path / "worker-workspace"
    workspace.mkdir()
    (workspace / "worker.stderr.log").write_text("closing", encoding="utf-8")
    original_rmtree = cancellable.shutil.rmtree
    attempts = 0

    def transient_rmtree(path: Path) -> None:
        nonlocal attempts
        attempts += 1
        if attempts < 3:
            raise PermissionError("worker handle is still closing")
        original_rmtree(path)

    monkeypatch.setattr(cancellable.shutil, "rmtree", transient_rmtree)

    assert cancellable._remove_workspace(workspace) is True
    assert attempts == 3
    assert not workspace.exists()


async def _invoke_excel_load(path: Path, timeout_seconds: float = 30.0) -> str:
    return await M._excel_load_tool(
        str(path),
        timeout_seconds=timeout_seconds,
    )


def test_tagged_result_codec_round_trips_excel_scalar_types() -> None:
    from datetime import date, datetime, time as datetime_time

    expected = {
        "datetime": datetime(2026, 8, 11, 12, 34, 56, 789000, fold=1),
        "date": date(2026, 8, 11),
        "time": datetime_time(12, 34, 56, 789000, fold=1),
        "bytes": b"\x00\x01docloupe\xff",
    }

    decoded = cancellable._result_json_loads(
        cancellable._result_json_bytes(expected)
    )

    assert decoded == expected
    assert decoded["datetime"].fold == 1
    assert decoded["time"].fold == 1


def test_public_excel_load_uses_async_cancellable_wrapper() -> None:
    tools = M.mcp._tool_manager._tools
    assert tools["excel_load"].fn is M._excel_load_tool
    assert M.excel_load is not M._excel_load_tool


def test_async_excel_load_preserves_datetime_and_reports_metrics(tmp_path: Path) -> None:
    from datetime import datetime

    source = tmp_path / "datetime.xlsx"
    expected = datetime(2026, 8, 11, 12, 34, 56)
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = expected
    workbook.save(source)
    workbook.close()

    load_result = anyio.run(_invoke_excel_load, source, 30.0)
    parsed_session_key = load_result.split("session_key='")[1].split("'")[0]
    session_key = M._resolve_session_key(parsed_session_key)
    data = M._sessions[session_key]

    assert data["sheets"][0]["rows"][0]["cells"][0]["v"] == expected
    metrics = data["_load_metrics"]
    assert metrics["worker_serialization_seconds"] >= 0.0
    assert metrics["worker_json_encode_seconds"] >= 0.0
    assert metrics["artifact_bytes"] > 0
    assert metrics["parent_json_decode_seconds"] >= 0.0
    assert metrics["parent_session_import_seconds"] >= 0.0
    assert metrics["total_tool_seconds"] >= metrics["parent_session_import_seconds"]
    assert "load_metrics=" in load_result
    M.excel_close(session_key)


def test_large_result_artifact_does_not_deadlock_and_returns_metrics() -> None:
    async def invoke() -> dict:
        return await run_named_operation(
            "_test_large_result",
            {"size": 2 * 1024 * 1024},
            timeout_seconds=20.0,
            operation_label="large-result-regression",
            return_metadata=True,
        )

    outcome = anyio.run(invoke)

    assert len(outcome["result"]["value"]) == 2 * 1024 * 1024
    assert outcome["metrics"]["worker"]["artifact_bytes"] > 2 * 1024 * 1024
    assert outcome["metrics"]["parent"]["artifact_bytes"] > 2 * 1024 * 1024
    assert outcome["metrics"]["parent"]["json_decode_seconds"] >= 0.0
    assert active_worker_pids() == ()


def test_result_artifact_limit_stops_worker_and_cleans_workspace(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from cancellable import WorkerProcessError

    workspaces = _capture_worker_workspaces(monkeypatch, tmp_path)
    monkeypatch.setenv("EXCEL_MCP_MAX_RESULT_BYTES", str(1024 * 1024))

    async def invoke() -> dict:
        return await run_named_operation(
            "_test_large_result",
            {"size": 2 * 1024 * 1024},
            timeout_seconds=20.0,
            operation_label="result-limit-regression",
        )

    with pytest.raises(WorkerProcessError) as caught:
        anyio.run(invoke)

    assert caught.value.details["code"] == "HEAVY_OPERATION_FAILED"
    assert caught.value.details["worker_stopped"] is True
    assert caught.value.details["workspace_removed"] is True
    assert caught.value.details["staging_removed"] is True
    assert len(workspaces) == 1
    assert not workspaces[0].exists()


@pytest.mark.parametrize("corruption", ["hash", "json"])
def test_invalid_result_artifact_is_rejected_and_cleaned(
    corruption: str,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from cancellable import WorkerProcessError

    workspaces = _capture_worker_workspaces(monkeypatch, tmp_path)
    original_read_status = cancellable._read_status

    def corrupt_status(path: Path) -> dict:
        metadata = original_read_status(path)
        if metadata.get("status") != "ok":
            return metadata
        if corruption == "hash":
            metadata["sha256"] = "0" * 64
        else:
            artifact = path.parent / cancellable._RESULT_FILE
            raw = b"\xffnot-json"
            artifact.write_bytes(raw)
            metadata["size"] = len(raw)
            metadata["sha256"] = cancellable._sha256(raw)
        return metadata

    monkeypatch.setattr(cancellable, "_read_status", corrupt_status)

    with pytest.raises(WorkerProcessError) as caught:
        anyio.run(
            _invoke_named_operation,
            "_test_sleep",
            {"seconds": 0.0},
            10.0,
            f"invalid-{corruption}-artifact",
        )

    assert caught.value.details["code"] == "HEAVY_OPERATION_RESULT_INVALID"
    assert caught.value.details["worker_stopped"] is True
    assert caught.value.details["workspace_removed"] is True
    assert caught.value.details["staging_removed"] is True
    assert len(workspaces) == 1
    assert not workspaces[0].exists()


def test_excel_load_worker_error_creates_no_session(tmp_path: Path) -> None:
    from cancellable import WorkerProcessError

    missing = tmp_path / "missing.xlsx"
    session_key = str(missing.resolve())

    with pytest.raises(WorkerProcessError):
        anyio.run(_invoke_excel_load, missing, 10.0)

    assert session_key not in M._sessions
    assert active_worker_pids() == ()


def test_excel_load_timeout_creates_no_session_and_follow_up_succeeds(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "timeout.xlsx"
    _write_simple_workbook(source)
    session_key = str(source.resolve())
    real_run_named_operation = run_named_operation

    async def slow_load(operation, payload, **kwargs):
        del operation, payload
        return await real_run_named_operation(
            "_test_sleep",
            {"seconds": 30.0},
            timeout_seconds=kwargs.get("timeout_seconds"),
            operation_label=kwargs.get("operation_label") or "excel_load",
        )

    monkeypatch.setattr(M, "run_named_operation", slow_load)

    with pytest.raises(WorkerProcessTimeout) as caught:
        anyio.run(_invoke_excel_load, source, 0.2)

    assert caught.value.details["worker_stopped"] is True
    assert caught.value.details["workspace_removed"] is True
    assert session_key not in M._sessions
    assert active_worker_pids() == ()
    follow_up = json.loads(M.excel_get_info(str(source)))
    assert follow_up["sheets"][0]["name"] == "Sheet"


def test_excel_load_cancellation_creates_no_session(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "cancel.xlsx"
    _write_simple_workbook(source)
    session_key = str(source.resolve())
    real_run_named_operation = run_named_operation
    outcome: dict[str, object] = {}

    async def slow_load(operation, payload, **kwargs):
        del operation, payload
        return await real_run_named_operation(
            "_test_sleep",
            {"seconds": 30.0},
            timeout_seconds=20.0,
            operation_label=kwargs.get("operation_label") or "excel_load",
        )

    async def invoke() -> None:
        try:
            await M._excel_load_tool(str(source), timeout_seconds=20.0)
        except WorkerProcessCancelled as exc:
            outcome["error"] = exc

    async def scenario() -> None:
        async with anyio.create_task_group() as task_group:
            task_group.start_soon(invoke)
            outcome["pid"] = await _wait_for_worker_pid()
            task_group.cancel_scope.cancel()

    monkeypatch.setattr(M, "run_named_operation", slow_load)
    anyio.run(scenario)

    error = outcome.get("error")
    assert isinstance(error, WorkerProcessCancelled)
    assert error.details["worker_stopped"] is True
    assert error.details["workspace_removed"] is True
    assert session_key not in M._sessions
    assert active_worker_pids() == ()
    assert not _process_is_running(outcome["pid"])
