from __future__ import annotations

import os
import subprocess
import sys


def snapshot_excel_pids() -> set[int]:
    if sys.platform != "win32":
        return set()
    return _list_excel_pids()


def cleanup_excel_spawned_since(before: set[int]) -> None:
    if sys.platform != "win32":
        return
    flag = (os.environ.get("MULTI_AGENT_SKIP_EXCEL_CLEANUP") or "").strip().lower()
    if flag in {"1", "true", "yes", "on"}:
        return
    for pid in sorted(_list_excel_pids() - set(before)):
        _taskkill(pid)


def _list_excel_pids() -> set[int]:
    result = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            "(Get-Process -Name EXCEL -ErrorAction SilentlyContinue) | ForEach-Object { $_.Id }",
        ],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    pids: set[int] = set()
    for line in (result.stdout or "").splitlines():
        value = line.strip()
        if value.isdecimal():
            pids.add(int(value))
    return pids


def _taskkill(pid: int) -> None:
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/T", "/F"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
