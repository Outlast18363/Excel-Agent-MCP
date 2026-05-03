"""Windows-only best-effort cleanup of Excel processes spawned during a runner.

Excel runs out-of-process (COM). Killing the Codex or MCP child process with
``taskkill`` does not always tear down ``EXCEL.EXE``, so automation can leak
many instances. :func:`snapshot_excel_pids` + :func:`cleanup_excel_spawned_since`
end only *new* EXCEL.EXE PIDs (present after the run but not in the snapshot).
"""

from __future__ import annotations

import logging
import os
import subprocess
import sys

_log = logging.getLogger(__name__)


def snapshot_excel_pids() -> set[int]:
    """Return the set of Windows PIDs for ``EXCEL.exe`` (empty on non-Windows)."""
    if sys.platform != "win32":
        return set()
    return _list_excel_pids_ps()


def cleanup_excel_spawned_since(before: set[int]) -> None:
    """Terminate every ``EXCEL.exe`` process whose PID was not in ``before``.

    Best-effort: never raises. Uses ``taskkill /F`` for processes that are still
    running so hung COM servers do not accumulate across Finch runs.

    Set environment variable ``MULTI_AGENT_SKIP_EXCEL_CLEANUP`` to a truthy
    value (``1`` / ``true`` / ``yes``) to skip this (e.g. if you need Excel
    left running next to a baseline run).
    """
    if sys.platform != "win32":
        return
    flag = (os.environ.get("MULTI_AGENT_SKIP_EXCEL_CLEANUP") or "").strip().lower()
    if flag in ("1", "true", "yes", "on"):
        return
    after = _list_excel_pids_ps()
    victims = after - set(before)
    if not victims:
        return
    _log.info("Reaping %d Excel process(es) not present before the run: %s", len(victims), sorted(victims))
    for pid in victims:
        _taskkill_tree(pid)


def _list_excel_pids_ps() -> set[int]:
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
        s = line.strip()
        if s.isdecimal():
            pids.add(int(s))
    return pids


def _taskkill_tree(pid: int) -> None:
    # /T ends child workbooks' helper processes if any; /F required when COM is stuck.
    completed = subprocess.run(
        ["taskkill", "/PID", str(pid), "/T", "/F"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if completed.returncode != 0 and completed.stderr and "not found" not in completed.stderr.lower():
        _log.debug("taskkill %s: %s", pid, completed.stderr.strip())
