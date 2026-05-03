"""Minimal Codex write-side-effect probe for the two-agent Worker config.

Run from the Excel-Agent-MCP repo root:

    python -m two_agent_framework.minimal_impl_write_probe

The probe launches `codex exec --json` through the same two-agent Worker
configuration (`build_codex_cmd("Worker", ...)`, optional Codex provider
overrides, MCP config, sandbox flags). It asks the worker to create one
implementation report, then records whether the file was actually created and
whether any tool command was attempted.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import build_codex_cmd, twowork_subprocess_env


def _json_dump(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _build_prompt(target: Path) -> str:
    return f"""\
You are the WORKER in a minimal side-effect probe.

Your first action must be a shell command or file-edit tool call that creates this exact file:
{target}

Write exactly this content to the file:
# Execution Todo

- [ ] A

# Implementation Report

## Completed Work
Probe file created.

After the file exists, verify it exists with a shell command. Do not inspect any workbook.
Do not finish with an assistant message until after the file has been created and verified.
"""


def _event_summary(events: list[dict[str, Any]]) -> dict[str, Any]:
    item_types: list[str] = []
    command_count = 0
    command_statuses: list[dict[str, Any]] = []
    final_message = ""
    turn_failed = False
    process_exit_codes: list[int] = []

    for event in events:
        event_type = event.get("type")
        if event_type == "turn.failed":
            turn_failed = True
        if event_type == "process.exited":
            code = event.get("return_code")
            if isinstance(code, int):
                process_exit_codes.append(code)
        if event_type != "item.completed" and event_type != "item.started":
            continue
        item = event.get("item") or {}
        item_type = item.get("type")
        if isinstance(item_type, str):
            item_types.append(item_type)
        if item_type == "command_execution":
            command_count += 1
            command_statuses.append(
                {
                    "id": item.get("id"),
                    "status": item.get("status"),
                    "exit_code": item.get("exit_code"),
                    "command": item.get("command"),
                }
            )
        if event_type == "item.completed" and item_type == "agent_message":
            text = item.get("text")
            if isinstance(text, str):
                final_message = text

    return {
        "item_types": item_types,
        "command_count": command_count,
        "command_statuses": command_statuses,
        "turn_failed": turn_failed,
        "process_exit_codes": process_exit_codes,
        "final_message": final_message,
    }


def run_probe(run_dir: Path, *, with_excel_mcp: bool, timeout_seconds: int) -> int:
    run_dir = run_dir.resolve()
    handover_dir = run_dir / "handover"
    handover_dir.mkdir(parents=True, exist_ok=True)

    target = handover_dir / "impl_report.md"
    trace_path = run_dir / "trace.jsonl"
    summary_path = run_dir / "summary.json"
    prompt_path = run_dir / "prompt.txt"

    target.unlink(missing_ok=True)
    trace_path.unlink(missing_ok=True)
    summary_path.unlink(missing_ok=True)

    prompt = _build_prompt(target)
    prompt_path.write_text(prompt, encoding="utf-8")

    cmd = build_codex_cmd("Worker", run_dir, with_excel_mcp=with_excel_mcp)
    env = twowork_subprocess_env()

    events: list[dict[str, Any]] = []
    raw_lines: list[str] = []

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        stdin=subprocess.PIPE,
        bufsize=0,
        env=env,
    )
    try:
        stdout, _ = proc.communicate(input=prompt.encode("utf-8"), timeout=timeout_seconds)
    except subprocess.TimeoutExpired:
        proc.kill()
        stdout, _ = proc.communicate()
        return_code = proc.wait()
        timed_out = True
    else:
        return_code = proc.returncode
        timed_out = False

    with trace_path.open("a", encoding="utf-8") as trace_file:
        for line in stdout.decode("utf-8", errors="replace").splitlines():
            if not line:
                continue
            raw_lines.append(line)
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                trace_file.write(json.dumps({"type": "raw", "line": line}, ensure_ascii=False) + "\n")
                continue
            if not isinstance(event, dict):
                trace_file.write(json.dumps({"type": "raw", "line": line}, ensure_ascii=False) + "\n")
                continue
            events.append(event)
            trace_file.write(json.dumps(event, ensure_ascii=False) + "\n")

    file_exists = target.exists()
    file_text = target.read_text(encoding="utf-8") if file_exists else ""
    summary = {
        "run_dir": str(run_dir),
        "target": str(target),
        "with_excel_mcp": with_excel_mcp,
        "return_code": return_code,
        "timed_out": timed_out,
        "file_exists": file_exists,
        "file_length": len(file_text),
        "event_summary": _event_summary(events),
        "trace": str(trace_path),
        "prompt": str(prompt_path),
        "raw_line_count": len(raw_lines),
    }
    _json_dump(summary_path, summary)

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if return_code == 0 and file_exists else 1


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--run-dir",
        type=Path,
        default=Path("probe_runs") / f"impl_write_{datetime.now():%Y%m%d_%H%M%S}",
        help="Directory where prompt.txt, trace.jsonl, summary.json, and handover/ are written.",
    )
    parser.add_argument(
        "--disable-excel-mcp",
        action="store_true",
        help="Disable excel-mcp. By default the probe keeps it enabled to match the Worker config.",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=300,
        help="Maximum time to wait for the Codex subprocess.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    return run_probe(
        run_dir=args.run_dir,
        with_excel_mcp=not args.disable_excel_mcp,
        timeout_seconds=args.timeout_seconds,
    )


if __name__ == "__main__":
    raise SystemExit(main())
