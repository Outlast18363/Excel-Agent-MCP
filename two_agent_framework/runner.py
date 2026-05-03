"""Entry point.

Example:
    python -m two_agent_framework.runner \\
        --task "Add a pivot summary sheet..." \\
        --workbooks /path/to/book.xlsx /path/to/supporting.pdf \\
        --run-dir trace_logs/run_20260419
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from .excel_lifecycle import cleanup_excel_spawned_since, snapshot_excel_pids
from .orchestrator import MAX_REDO, MAX_RESET, Orchestrator, unlink_or_truncate_run_file


def _derive_workbook_dir(workbooks: list[Path]) -> Path:
    if not workbooks:
        raise ValueError("either --workbook-dir or --workbooks is required")
    parents = {path.parent for path in workbooks}
    if len(parents) != 1:
        raise ValueError("--workbooks must all be staged in the same directory")
    return next(iter(parents))


def _parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(prog="two_agent_framework.runner")
    task_group = p.add_mutually_exclusive_group(required=True)
    task_group.add_argument("--task", help="natural-language task for the agent team")
    task_group.add_argument(
        "--task-json-path",
        type=Path,
        help="Path to a JSON file (e.g. Finch task record); full serialized JSON is passed to agents.",
    )
    p.add_argument(
        "--workbooks",
        nargs="+",
        type=Path,
        default=None,
        help="Task workbook/source file paths staged inside the run workbook directory. "
             "The framework snapshots this full set and restores the folder from it on reset.",
    )
    p.add_argument(
        "--workbook-dir",
        type=Path,
        default=None,
        help="Path to the staged workbook/source-file directory. Preferred over --workbooks.",
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
    p.add_argument(
        "--disable-excel-mcp",
        action="store_true",
        help="Do not configure the excel-mcp server for agent subprocesses.",
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
    unlink_or_truncate_run_file(run_dir / "trace.jsonl")
    unlink_or_truncate_run_file(run_dir / "wrapper_summary.json")

    workbook_dir: Path | None = None
    if args.workbook_dir is not None:
        workbook_dir = args.workbook_dir.resolve()
        if not workbook_dir.exists():
            print(f"workbook directory not found: {workbook_dir}", file=sys.stderr)
            return 2
        if not workbook_dir.is_dir():
            print(f"workbook directory is not a directory: {workbook_dir}", file=sys.stderr)
            return 2
        if not any(path.is_file() or path.is_symlink() for path in workbook_dir.iterdir()):
            print(f"workbook directory is empty: {workbook_dir}", file=sys.stderr)
            return 2
    else:
        workbooks = [path.resolve() for path in (args.workbooks or [])]
        missing = [path for path in workbooks if not path.exists()]
        if missing:
            print(f"workbook not found: {missing[0]}", file=sys.stderr)
            return 2
        try:
            workbook_dir = _derive_workbook_dir(workbooks)
        except ValueError as exc:
            print(str(exc), file=sys.stderr)
            return 2

    # When --task-id is omitted, fall back to the run_dir basename so the run
    # still has a stable identifier in logs and summaries.
    task_id = args.task_id or run_dir.name

    if args.task_json_path is not None:
        task_path = args.task_json_path.expanduser().resolve()
        if not task_path.is_file():
            print(f"task json not found: {task_path}", file=sys.stderr)
            return 2
        try:
            task_obj = json.loads(task_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            print(f"invalid JSON in {task_path}: {exc}", file=sys.stderr)
            return 2
        if not isinstance(task_obj, dict):
            print(f"task JSON must be an object: {task_path}", file=sys.stderr)
            return 2
        task_text = json.dumps(task_obj, indent=2, ensure_ascii=False)
    else:
        task_text = str(args.task)

    excel_pids_before = snapshot_excel_pids()
    try:
        result = Orchestrator(
            task=task_text,
            workbook_dir=workbook_dir,
            empty_workbook_created=args.empty_workbook_created,
            run_dir=run_dir,
            task_id=task_id,
            with_excel_mcp=not args.disable_excel_mcp,
        ).run()
    finally:
        cleanup_excel_spawned_since(excel_pids_before)

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
