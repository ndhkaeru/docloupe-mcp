from __future__ import annotations

import json
import os
import signal
import subprocess
import sys
import threading
import time
from pathlib import Path

if os.name == "nt":
    import ctypes
    from ctypes import wintypes

    _KERNEL32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    _STILL_ACTIVE = 259
    _KERNEL32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    _KERNEL32.OpenProcess.restype = wintypes.HANDLE
    _KERNEL32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    _KERNEL32.GetExitCodeProcess.restype = wintypes.BOOL
    _KERNEL32.CloseHandle.argtypes = [wintypes.HANDLE]
    _KERNEL32.CloseHandle.restype = wintypes.BOOL

import anyio
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "servers"))

from process_lifecycle import (
    ManagedProcessCancelled,
    ManagedProcessTimeout,
    active_managed_process_pids,
    run_cancellable_in_thread,
    run_managed_process,
)

HELPER = ROOT / "tests" / "fixtures" / "process_tree_helper.py"


def _pid_exists(process_id: int) -> bool:
    if os.name == "nt":
        handle = _KERNEL32.OpenProcess(
            _PROCESS_QUERY_LIMITED_INFORMATION,
            False,
            process_id,
        )
        if not handle:
            return False
        try:
            exit_code = wintypes.DWORD()
            if not _KERNEL32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                return False
            return exit_code.value == _STILL_ACTIVE
        finally:
            _KERNEL32.CloseHandle(handle)
    try:
        os.kill(process_id, 0)
    except OSError:
        return False
    return True


def _wait_for_pid_file(path: Path, timeout_seconds: float = 5.0) -> dict[str, int]:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if path.is_file():
            return json.loads(path.read_text(encoding="utf-8"))
        time.sleep(0.025)
    raise AssertionError(f"Timed out waiting for PID file: {path}")


def _wait_for_processes_to_stop(process_ids: list[int], timeout_seconds: float = 5.0) -> bool:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if not any(_pid_exists(process_id) for process_id in process_ids):
            return True
        time.sleep(0.05)
    return not any(_pid_exists(process_id) for process_id in process_ids)


def _force_stop(process_ids: list[int]) -> None:
    for process_id in process_ids:
        if not _pid_exists(process_id):
            continue
        if os.name == "nt":
            subprocess.run(
                ["taskkill.exe", "/PID", str(process_id), "/T", "/F"],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        else:
            try:
                os.kill(process_id, signal.SIGKILL)
            except ProcessLookupError:
                pass


def test_managed_process_returns_output_and_releases_registry():
    result = run_managed_process(
        [sys.executable, "-c", "print('managed-ok')"],
        timeout_seconds=5.0,
    )

    assert result.returncode == 0
    assert result.stdout.strip() == "managed-ok"
    assert result.process_tree_stopped is True
    assert active_managed_process_pids() == ()


def test_timeout_stops_parent_and_descendant(tmp_path):
    pid_path = tmp_path / "timeout-pids.json"
    process_ids: list[int] = []
    try:
        with pytest.raises(ManagedProcessTimeout) as captured:
            run_managed_process(
                [sys.executable, str(HELPER), "parent", str(pid_path)],
                timeout_seconds=1.0,
                terminate_grace_seconds=0.1,
                kill_grace_seconds=3.0,
            )
        process_ids = list(_wait_for_pid_file(pid_path).values())

        assert captured.value.process_tree_stopped is True
        assert _wait_for_processes_to_stop(process_ids)
        assert active_managed_process_pids() == ()
    finally:
        _force_stop(process_ids)


def test_cancel_event_stops_parent_and_descendant(tmp_path):
    pid_path = tmp_path / "cancel-pids.json"
    cancel_event = threading.Event()
    failures: list[BaseException] = []
    process_ids: list[int] = []

    def invoke() -> None:
        try:
            run_managed_process(
                [sys.executable, str(HELPER), "parent", str(pid_path)],
                timeout_seconds=30.0,
                cancel_event=cancel_event,
                terminate_grace_seconds=0.1,
                kill_grace_seconds=3.0,
            )
        except BaseException as exc:
            failures.append(exc)

    worker = threading.Thread(target=invoke, daemon=True)
    worker.start()
    try:
        process_ids = list(_wait_for_pid_file(pid_path).values())
        cancel_event.set()
        worker.join(timeout=10.0)

        assert not worker.is_alive()
        assert len(failures) == 1
        assert isinstance(failures[0], ManagedProcessCancelled)
        assert failures[0].process_tree_stopped is True
        assert _wait_for_processes_to_stop(process_ids)
        assert active_managed_process_pids() == ()
    finally:
        cancel_event.set()
        worker.join(timeout=3.0)
        _force_stop(process_ids)


def test_async_cancellation_waits_for_thread_cleanup():
    started = threading.Event()
    stopped = threading.Event()

    def cancellable_work(cancel_event: threading.Event) -> None:
        started.set()
        cancel_event.wait(5.0)
        stopped.set()
        raise ManagedProcessCancelled(("fake",), process_tree_stopped=True)

    async def scenario() -> None:
        with anyio.move_on_after(0.1) as cancel_scope:
            await run_cancellable_in_thread(cancellable_work)
        assert cancel_scope.cancel_called is True

    anyio.run(scenario)

    assert started.is_set()
    assert stopped.is_set()
    assert active_managed_process_pids() == ()
