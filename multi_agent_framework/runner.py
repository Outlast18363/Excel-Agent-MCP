"""Entry point.

Example:
    python -m multi_agent_framework.runner \\
        --task "Add a pivot summary sheet..." \\
        --workbook /path/to/book.xlsx \\
        --run-dir trace_logs/run_20260419
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from .orchestrator import MAX_REDO, MAX_RESET, Orchestrator


def _parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(prog="multi_agent_framework.runner")
    p.add_argument("--task", required=True, help="natural-language task for the agent team")
    p.add_argument("--workbook", required=True, type=Path, help="path to the .xlsx to operate on")
    p.add_argument(
        "--run-dir",
        type=Path,
        default=None,
        help="directory for handover/, snapshot.xlsx, trace.jsonl (default: trace_logs/run_<ts>)",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(sys.argv[1:] if argv is None else argv)

    run_dir = args.run_dir or Path("trace_logs") / f"run_{datetime.now():%Y%m%d_%H%M%S}"
    run_dir.mkdir(parents=True, exist_ok=True)

    if not args.workbook.exists():
        print(f"workbook not found: {args.workbook}", file=sys.stderr)
        return 2

    result = Orchestrator(task=args.task, workbook=args.workbook, run_dir=run_dir).run()

    summary = {
        "verdict": result.verdict,
        "iterations": result.iterations,
        "redo_count": result.redo_count,
        "reset_count": result.reset_count,
        "max_redo": MAX_REDO,
        "max_reset": MAX_RESET,
        "trace": str(result.trace_path),
        "run_dir": str(result.run_dir),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if result.verdict == "success" else 1


if __name__ == "__main__":
    raise SystemExit(main())
