"""Bounded external-process execution with process-tree cleanup."""
from __future__ import annotations

import os
import shutil
import signal
import subprocess
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence, TypeVar

_POLL_SECONDS = 0.05
_DEFAULT_TERMINATE_GRACE_SECONDS = 1.0
_DEFAULT_KILL_GRACE_SECONDS = 3.0
_DEFAULT_POST_EXIT_GRACE_SECONDS = 0.25
_ACTIVE_PROCESS_LOCK = threading.Lock()
_ACTIVE_PROCESS_PIDS: set[int] = set()
_ResultT = TypeVar("_ResultT")


@dataclass(frozen=True)
class ManagedProcessResult:
    args: tuple[str, ...]
    returncode: int
    stdout: str | bytes | None
    stderr: str | bytes | None
    elapsed_seconds: float
    process_tree_stopped: bool


class ManagedProcessCancelled(RuntimeError):
    def __init__(
        self,
        command: Sequence[str],
        *,
        process_tree_stopped: bool,
        stdout: str | bytes | None = None,
        stderr: str | bytes | None = None,
    ) -> None:
        self.command = tuple(str(item) for item in command)
        self.process_tree_stopped = process_tree_stopped
        self.stdout = stdout
        self.stderr = stderr
        super().__init__(
            f"External process was cancelled; process_tree_stopped={process_tree_stopped}."
        )


class ManagedProcessTimeout(subprocess.TimeoutExpired):
    def __init__(
        self,
        command: Sequence[str],
        timeout_seconds: float,
        *,
        process_tree_stopped: bool,
        stdout: str | bytes | None = None,
        stderr: str | bytes | None = None,
    ) -> None:
        super().__init__(
            tuple(str(item) for item in command),
            timeout_seconds,
            output=stdout,
            stderr=stderr,
        )
        self.process_tree_stopped = process_tree_stopped


class ManagedProcessCleanupError(RuntimeError):
    def __init__(self, process_id: int | None, message: str) -> None:
        self.process_id = process_id
        super().__init__(message)


if os.name == "nt":
    import ctypes
    from ctypes import wintypes

    _KERNEL32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000
    _JOB_OBJECT_BASIC_ACCOUNTING_INFORMATION = 1
    _JOB_OBJECT_EXTENDED_LIMIT_INFORMATION = 9

    class _IOCounters(ctypes.Structure):
        _fields_ = [
            ("ReadOperationCount", ctypes.c_ulonglong),
            ("WriteOperationCount", ctypes.c_ulonglong),
            ("OtherOperationCount", ctypes.c_ulonglong),
            ("ReadTransferCount", ctypes.c_ulonglong),
            ("WriteTransferCount", ctypes.c_ulonglong),
            ("OtherTransferCount", ctypes.c_ulonglong),
        ]

    class _JobBasicLimitInformation(ctypes.Structure):
        _fields_ = [
            ("PerProcessUserTimeLimit", ctypes.c_longlong),
            ("PerJobUserTimeLimit", ctypes.c_longlong),
            ("LimitFlags", wintypes.DWORD),
            ("MinimumWorkingSetSize", ctypes.c_size_t),
            ("MaximumWorkingSetSize", ctypes.c_size_t),
            ("ActiveProcessLimit", wintypes.DWORD),
            ("Affinity", ctypes.c_size_t),
            ("PriorityClass", wintypes.DWORD),
            ("SchedulingClass", wintypes.DWORD),
        ]

    class _JobExtendedLimitInformation(ctypes.Structure):
        _fields_ = [
            ("BasicLimitInformation", _JobBasicLimitInformation),
            ("IoInfo", _IOCounters),
            ("ProcessMemoryLimit", ctypes.c_size_t),
            ("JobMemoryLimit", ctypes.c_size_t),
            ("PeakProcessMemoryUsed", ctypes.c_size_t),
            ("PeakJobMemoryUsed", ctypes.c_size_t),
        ]

    class _JobBasicAccountingInformation(ctypes.Structure):
        _fields_ = [
            ("TotalUserTime", ctypes.c_longlong),
            ("TotalKernelTime", ctypes.c_longlong),
            ("ThisPeriodTotalUserTime", ctypes.c_longlong),
            ("ThisPeriodTotalKernelTime", ctypes.c_longlong),
            ("TotalPageFaultCount", wintypes.DWORD),
            ("TotalProcesses", wintypes.DWORD),
            ("ActiveProcesses", wintypes.DWORD),
            ("TotalTerminatedProcesses", wintypes.DWORD),
        ]

    _KERNEL32.CreateJobObjectW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR]
    _KERNEL32.CreateJobObjectW.restype = wintypes.HANDLE
    _KERNEL32.SetInformationJobObject.argtypes = [
        wintypes.HANDLE,
        ctypes.c_int,
        ctypes.c_void_p,
        wintypes.DWORD,
    ]
    _KERNEL32.SetInformationJobObject.restype = wintypes.BOOL
    _KERNEL32.AssignProcessToJobObject.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
    _KERNEL32.AssignProcessToJobObject.restype = wintypes.BOOL
    _KERNEL32.QueryInformationJobObject.argtypes = [
        wintypes.HANDLE,
        ctypes.c_int,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.c_void_p,
    ]
    _KERNEL32.QueryInformationJobObject.restype = wintypes.BOOL
    _KERNEL32.TerminateJobObject.argtypes = [wintypes.HANDLE, wintypes.UINT]
    _KERNEL32.TerminateJobObject.restype = wintypes.BOOL
    _KERNEL32.CloseHandle.argtypes = [wintypes.HANDLE]
    _KERNEL32.CloseHandle.restype = wintypes.BOOL


class _WindowsJob:
    def __init__(self, process: subprocess.Popen[Any]) -> None:
        if os.name != "nt":
            raise OSError("Windows Job Objects are unavailable on this platform.")
        handle = _KERNEL32.CreateJobObjectW(None, None)
        if not handle:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            limits = _JobExtendedLimitInformation()
            limits.BasicLimitInformation.LimitFlags = _JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            if not _KERNEL32.SetInformationJobObject(
                handle,
                _JOB_OBJECT_EXTENDED_LIMIT_INFORMATION,
                ctypes.byref(limits),
                ctypes.sizeof(limits),
            ):
                raise ctypes.WinError(ctypes.get_last_error())
            process_handle = wintypes.HANDLE(int(process._handle))
            if not _KERNEL32.AssignProcessToJobObject(handle, process_handle):
                raise ctypes.WinError(ctypes.get_last_error())
        except BaseException:
            _KERNEL32.CloseHandle(handle)
            raise
        self._handle = handle

    def active_processes(self) -> int | None:
        if self._handle is None:
            return 0
        accounting = _JobBasicAccountingInformation()
        if not _KERNEL32.QueryInformationJobObject(
            self._handle,
            _JOB_OBJECT_BASIC_ACCOUNTING_INFORMATION,
            ctypes.byref(accounting),
            ctypes.sizeof(accounting),
            None,
        ):
            return None
        return int(accounting.ActiveProcesses)

    def terminate(self, exit_code: int = 1) -> bool:
        if self._handle is None:
            return True
        return bool(_KERNEL32.TerminateJobObject(self._handle, exit_code))

    def close(self) -> None:
        if self._handle is not None:
            _KERNEL32.CloseHandle(self._handle)
            self._handle = None


class _ProcessTree:
    def __init__(self, process: subprocess.Popen[Any]) -> None:
        self.process = process
        self.process_id = process.pid
        self.windows_job: _WindowsJob | None = None
        self.windows_job_error: str | None = None
        self.taskkill_succeeded = False
        if os.name == "nt":
            try:
                self.windows_job = _WindowsJob(process)
            except OSError as exc:
                self.windows_job_error = str(exc)

    def is_alive(self) -> bool:
        if self.process_id is None:
            return False
        if os.name == "nt":
            if self.windows_job is not None:
                active_processes = self.windows_job.active_processes()
                if active_processes is not None:
                    return active_processes > 0
            return self.process.poll() is None and not self.taskkill_succeeded
        try:
            os.killpg(self.process_id, 0)
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        return True

    def terminate(self) -> None:
        if self.process_id is None:
            return
        if os.name == "nt":
            self.taskkill_succeeded = _taskkill_tree(self.process_id, force=False)
            if self.taskkill_succeeded:
                return
            ctrl_break_event = getattr(signal, "CTRL_BREAK_EVENT", None)
            if ctrl_break_event is not None and self.process.poll() is None:
                try:
                    self.process.send_signal(ctrl_break_event)
                except OSError:
                    pass
            return
        try:
            os.killpg(self.process_id, signal.SIGTERM)
        except ProcessLookupError:
            pass

    def kill(self) -> None:
        if self.process_id is None:
            return
        if os.name == "nt":
            self.taskkill_succeeded = _taskkill_tree(self.process_id, force=True)
            if self.windows_job is not None:
                self.windows_job.terminate(1)
            if not self.taskkill_succeeded and self.process.poll() is None:
                try:
                    self.process.kill()
                except OSError:
                    pass
            return
        try:
            os.killpg(self.process_id, signal.SIGKILL)
        except ProcessLookupError:
            pass

    def wait_for_exit(self, timeout_seconds: float) -> bool:
        deadline = time.monotonic() + max(timeout_seconds, 0.0)
        while self.is_alive():
            if time.monotonic() >= deadline:
                return False
            time.sleep(_POLL_SECONDS)
        return True

    def close(self) -> None:
        if self.windows_job is not None:
            self.windows_job.close()


def _taskkill_tree(process_id: int, *, force: bool) -> bool:
    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    taskkill_path = system_root / "System32" / "taskkill.exe"
    command = str(taskkill_path) if taskkill_path.is_file() else "taskkill.exe"
    arguments = [command, "/PID", str(process_id), "/T"]
    if force:
        arguments.append("/F")
    try:
        completed = subprocess.run(
            arguments,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return completed.returncode == 0


def active_managed_process_pids() -> tuple[int, ...]:
    with _ACTIVE_PROCESS_LOCK:
        return tuple(sorted(_ACTIVE_PROCESS_PIDS))


def _register_process(process_id: int | None) -> None:
    if process_id is None:
        return
    with _ACTIVE_PROCESS_LOCK:
        _ACTIVE_PROCESS_PIDS.add(process_id)


def _unregister_process(process_id: int | None) -> None:
    if process_id is None:
        return
    with _ACTIVE_PROCESS_LOCK:
        _ACTIVE_PROCESS_PIDS.discard(process_id)


def _stop_process_tree(
    process: subprocess.Popen[Any],
    tree: _ProcessTree,
    *,
    terminate_grace_seconds: float,
    kill_grace_seconds: float,
) -> bool:
    if not tree.is_alive():
        return True
    tree.terminate()
    if tree.wait_for_exit(terminate_grace_seconds):
        return True
    tree.kill()
    try:
        process.wait(timeout=max(kill_grace_seconds, 0.0))
    except subprocess.TimeoutExpired:
        pass
    return tree.wait_for_exit(kill_grace_seconds)


def _drain_process(
    process: subprocess.Popen[Any],
    timeout_seconds: float,
) -> tuple[str | bytes | None, str | bytes | None]:
    try:
        return process.communicate(timeout=max(timeout_seconds, 0.0))
    except subprocess.TimeoutExpired as exc:
        return exc.output, exc.stderr


def run_managed_process(
    command: Sequence[str | os.PathLike[str]],
    *,
    timeout_seconds: float | None,
    cancel_event: threading.Event | None = None,
    cwd: str | os.PathLike[str] | None = None,
    env: Mapping[str, str] | None = None,
    capture_output: bool = True,
    text: bool = True,
    encoding: str | None = None,
    errors: str | None = None,
    check: bool = False,
    terminate_grace_seconds: float = _DEFAULT_TERMINATE_GRACE_SECONDS,
    kill_grace_seconds: float = _DEFAULT_KILL_GRACE_SECONDS,
    post_exit_grace_seconds: float = _DEFAULT_POST_EXIT_GRACE_SECONDS,
    windows_hide: bool = True,
) -> ManagedProcessResult:
    normalized_command = tuple(os.fspath(item) for item in command)
    if not normalized_command:
        raise ValueError("command must not be empty")
    if timeout_seconds is not None and timeout_seconds <= 0:
        raise ValueError("timeout_seconds must be positive or None")
    if cancel_event is not None and cancel_event.is_set():
        raise ManagedProcessCancelled(normalized_command, process_tree_stopped=True)

    popen_arguments: dict[str, Any] = {
        "stdin": subprocess.DEVNULL,
        "stdout": subprocess.PIPE if capture_output else None,
        "stderr": subprocess.PIPE if capture_output else None,
        "cwd": cwd,
        "env": dict(env) if env is not None else None,
        "text": text,
        "close_fds": True,
    }
    if encoding is not None:
        popen_arguments["encoding"] = encoding
    if errors is not None:
        popen_arguments["errors"] = errors
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        if windows_hide:
            creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
        popen_arguments["creationflags"] = creationflags
    else:
        popen_arguments["start_new_session"] = True

    started_at = time.monotonic()
    process = subprocess.Popen(normalized_command, **popen_arguments)
    tree = _ProcessTree(process)
    _register_process(process.pid)
    deadline = None if timeout_seconds is None else started_at + timeout_seconds
    stdout: str | bytes | None = None
    stderr: str | bytes | None = None

    try:
        while True:
            if cancel_event is not None and cancel_event.is_set():
                process_tree_stopped = _stop_process_tree(
                    process,
                    tree,
                    terminate_grace_seconds=terminate_grace_seconds,
                    kill_grace_seconds=kill_grace_seconds,
                )
                stdout, stderr = _drain_process(process, kill_grace_seconds)
                raise ManagedProcessCancelled(
                    normalized_command,
                    process_tree_stopped=process_tree_stopped,
                    stdout=stdout,
                    stderr=stderr,
                )

            remaining = None if deadline is None else deadline - time.monotonic()
            if remaining is not None and remaining <= 0:
                process_tree_stopped = _stop_process_tree(
                    process,
                    tree,
                    terminate_grace_seconds=terminate_grace_seconds,
                    kill_grace_seconds=kill_grace_seconds,
                )
                stdout, stderr = _drain_process(process, kill_grace_seconds)
                raise ManagedProcessTimeout(
                    normalized_command,
                    timeout_seconds,
                    process_tree_stopped=process_tree_stopped,
                    stdout=stdout,
                    stderr=stderr,
                )

            wait_seconds = _POLL_SECONDS if remaining is None else min(_POLL_SECONDS, remaining)
            try:
                stdout, stderr = process.communicate(timeout=max(wait_seconds, 0.001))
                break
            except subprocess.TimeoutExpired:
                continue

        process_tree_stopped = tree.wait_for_exit(post_exit_grace_seconds)
        if not process_tree_stopped:
            process_tree_stopped = _stop_process_tree(
                process,
                tree,
                terminate_grace_seconds=terminate_grace_seconds,
                kill_grace_seconds=kill_grace_seconds,
            )
        if not process_tree_stopped:
            raise ManagedProcessCleanupError(
                process.pid,
                f"External process tree for PID {process.pid} did not stop after command exit.",
            )

        result = ManagedProcessResult(
            args=normalized_command,
            returncode=int(process.returncode or 0),
            stdout=stdout,
            stderr=stderr,
            elapsed_seconds=round(time.monotonic() - started_at, 6),
            process_tree_stopped=True,
        )
        if check and result.returncode != 0:
            raise subprocess.CalledProcessError(
                result.returncode,
                result.args,
                output=result.stdout,
                stderr=result.stderr,
            )
        return result
    except BaseException:
        if tree.is_alive():
            _stop_process_tree(
                process,
                tree,
                terminate_grace_seconds=terminate_grace_seconds,
                kill_grace_seconds=kill_grace_seconds,
            )
        raise
    finally:
        tree.close()
        _unregister_process(process.pid)


async def run_cancellable_in_thread(
    function: Callable[[threading.Event], _ResultT],
) -> _ResultT:
    import anyio

    cancel_event = threading.Event()
    completed_event = threading.Event()

    def invoke() -> _ResultT:
        try:
            return function(cancel_event)
        finally:
            completed_event.set()

    try:
        return await anyio.to_thread.run_sync(invoke, abandon_on_cancel=True)
    except anyio.get_cancelled_exc_class():
        cancel_event.set()
        with anyio.CancelScope(shield=True):
            while not completed_event.is_set():
                await anyio.sleep(_POLL_SECONDS)
        raise


def remove_path_with_retries(path: str | os.PathLike[str]) -> bool:
    target = Path(path)
    for delay_seconds in (0.0, 0.05, 0.2, 0.5, 1.0):
        if delay_seconds:
            time.sleep(delay_seconds)
        try:
            if target.is_dir() and not target.is_symlink():
                shutil.rmtree(target)
            else:
                target.unlink(missing_ok=True)
        except OSError:
            continue
        return not target.exists()
    return not target.exists()
