"""Entry point.

Example:
    python -m multi_agent_framework.runner \\
        --task "Add a pivot summary sheet..." \\
        --workbooks /path/to/book.xlsx /path/to/supporting.pdf \\
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
    p.add_argument(
        "--workbooks",
        required=True,
        nargs="+",
        type=Path,
        help="Task workbook/source file paths staged inside the run workbook directory. "
             "The framework snapshots this full set and restores the folder from it on reset.",
    )
    p.add_argument(
        "--run-dir",
        type=Path,
        default=None,
        help="directory for handover/, snapshots/, trace.jsonl (default: trace_logs/run_<ts>)",
    )
    p.add_argument(
        "--task-id",
        default=None,
        help="Identifier used to sign the orchestrator's workbook copy "
             "inside final_result/ (e.g. Finch dataset id). "
             "Defaults to run_dir basename.",
    )
    p.add_argument(
        "--empty-workbook-created",
        action="store_true",
        help="Indicates the wrapper synthesized a blank .xlsx because the task had no source workbook.",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(sys.argv[1:] if argv is None else argv)

    run_dir = args.run_dir or Path("trace_logs") / f"run_{datetime.now():%Y%m%d_%H%M%S}"
    run_dir.mkdir(parents=True, exist_ok=True)

    # A re-run in the same run_dir must start clean: otherwise EventBus (opened
    # in append mode) would interleave new events into the previous run's
    # trace.jsonl, and a stale wrapper_summary.json from a prior run could be
    # mistaken for this run's result if this run crashes before its wrapper
    # rewrites it.
    (run_dir / "trace.jsonl").unlink(missing_ok=True)
    (run_dir / "wrapper_summary.json").unlink(missing_ok=True)

    missing = [path for path in args.workbooks if not path.exists()]
    if missing:
        print(f"workbook not found: {missing[0]}", file=sys.stderr)
        return 2

    # When --task-id is omitted, fall back to the run_dir basename so the run
    # still has a stable identifier in logs and summaries.
    task_id = args.task_id or run_dir.name

    result = Orchestrator(
        task=args.task,
        workbooks=args.workbooks,
        empty_workbook_created=args.empty_workbook_created,
        run_dir=run_dir,
        task_id=task_id,
    ).run()

    summary = {
        "verdict": result.verdict,
        "iterations": result.iterations,
        "redo_count": result.redo_count,
        "reset_count": result.reset_count,
        "max_redo": MAX_REDO,
        "max_reset": MAX_RESET,
        "trace": str(result.trace_path),
        "run_dir": str(result.run_dir),
        "final_dir": str(result.final_dir),
        "usage_by_agent": result.usage_by_agent or {},
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if result.verdict == "success" else 1


if __name__ == "__main__":
    raise SystemExit(main())
