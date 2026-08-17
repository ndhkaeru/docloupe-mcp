from __future__ import annotations

import json
import os
import signal
import subprocess
import sys
import time
from pathlib import Path


def _ignore_shutdown_signals() -> None:
    def ignore(_signum, _frame):
        return None

    for signal_name in ("SIGINT", "SIGTERM"):
        signal_value = getattr(signal, signal_name, None)
        if signal_value is not None:
            signal.signal(signal_value, ignore)


def _run_grandchild() -> int:
    _ignore_shutdown_signals()
    while True:
        time.sleep(1)


def _run_parent(pid_path: Path) -> int:
    _ignore_shutdown_signals()
    grandchild = subprocess.Popen(
        [sys.executable, __file__, "grandchild"],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
    )
    temporary_path = pid_path.with_suffix(".tmp")
    temporary_path.write_text(
        json.dumps({"parent": os.getpid(), "grandchild": grandchild.pid}),
        encoding="utf-8",
    )
    temporary_path.replace(pid_path)
    while True:
        time.sleep(1)


def main() -> int:
    if len(sys.argv) < 2:
        return 2
    if sys.argv[1] == "grandchild":
        return _run_grandchild()
    if sys.argv[1] == "parent" and len(sys.argv) == 3:
        return _run_parent(Path(sys.argv[2]))
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
