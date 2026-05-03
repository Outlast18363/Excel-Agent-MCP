from __future__ import annotations

import shutil
import subprocess
import sys
import time
from pathlib import Path


def cleanup_before_run(run_dir: Path, *, workbook_dir: Path | None = None, remove_run_contents: bool) -> None:
    """Best-effort cleanup for reusing a run directory.

    Process cleanup is scoped to command lines containing the run/workbook path,
    so unrelated Claude Code sessions are left alone. `remove_run_contents=True`
    is for the outer baseline wrapper before it stages fresh task inputs.
    """
    run_dir = Path(run_dir).resolve()
    paths = [run_dir]
    if workbook_dir is not None:
        paths.append(Path(workbook_dir).resolve())

    terminate_processes_referencing(paths)
    kill_all_excel_processes()
    if remove_run_contents and run_dir.exists():
        _clear_directory_contents(run_dir)


def terminate_processes_referencing(paths: list[Path]) -> list[int]:
    if sys.platform != "win32":
        return []
    needles = [str(path).lower() for path in paths]
    victims = _find_processes_referencing(needles)
    for pid in victims:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    if victims:
        time.sleep(0.5)
    return victims


def kill_all_excel_processes() -> None:
    if sys.platform != "win32":
        return
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            "Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force",
        ],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def _find_processes_referencing(needles: list[str]) -> list[int]:
    script = (
        "Get-CimInstance Win32_Process | "
        "Select-Object ProcessId,Name,CommandLine | "
        "ConvertTo-Json -Compress"
    )
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", script],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    try:
        import json

        payload = json.loads(result.stdout or "[]")
    except Exception:
        return []
    if isinstance(payload, dict):
        payload = [payload]
    victims: list[int] = []
    current_pid = None
    try:
        import os

        current_pid = os.getpid()
    except Exception:
        pass
    for row in payload if isinstance(payload, list) else []:
        if not isinstance(row, dict):
            continue
        name = str(row.get("Name") or "").lower()
        command_line = str(row.get("CommandLine") or "").lower()
        if not command_line:
            continue
        if name == "claude.exe":
            pass
        elif name == "node.exe" and "claude" in command_line:
            pass
        elif name in {"python.exe", "pythonw.exe"} and "\\claude_deepseek_two_agent\\main.py" in command_line:
            pass
        else:
            continue
        if not any(needle and needle in command_line for needle in needles):
            continue
        try:
            pid = int(row.get("ProcessId"))
        except (TypeError, ValueError):
            continue
        if current_pid is not None and pid == current_pid:
            continue
        victims.append(pid)
    return victims


def clear_run_owned_dirs(run_dir: Path) -> None:
    for name in ("handover", "snapshots", "screenshots", "final_result", "mcp_configs"):
        path = Path(run_dir) / name
        if path.exists():
            _remove_path(path)


def _clear_directory_contents(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for child in path.iterdir():
        _remove_path(child)


def _remove_path(path: Path) -> None:
    for attempt in range(3):
        try:
            if path.is_file() or path.is_symlink():
                path.unlink()
            else:
                shutil.rmtree(path)
            return
        except FileNotFoundError:
            return
        except PermissionError:
            if attempt == 2:
                raise
            time.sleep(0.5)
