from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from claude_invocation import validate_child_env
from excel_cleanup import cleanup_excel_spawned_since, snapshot_excel_pids
from orchestrator import MAX_REDO, MAX_RESET, Orchestrator, unlink_or_truncate_run_file
from run_cleanup import cleanup_before_run, clear_run_owned_dirs


def _derive_workbook_dir(workbooks: list[Path]) -> Path:
    if not workbooks:
        raise ValueError("either --workbook-dir or --workbooks is required")
    parents = {path.parent for path in workbooks}
    if len(parents) != 1:
        raise ValueError("--workbooks must all be staged in the same directory")
    return next(iter(parents))


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog="claude_deepseek_two_agent")
    task_group = parser.add_mutually_exclusive_group(required=True)
    task_group.add_argument("--task", help="natural-language task for the agent team")
    task_group.add_argument(
        "--task-json-path",
        type=Path,
        help="Path to a JSON file; full serialized JSON is passed to agents.",
    )
    parser.add_argument(
        "--workbooks",
        nargs="+",
        type=Path,
        default=None,
        help="Task workbook/source file paths staged inside the run workbook directory.",
    )
    parser.add_argument(
        "--workbook-dir",
        type=Path,
        default=None,
        help="Path to the staged workbook/source-file directory. Preferred over --workbooks.",
    )
    parser.add_argument(
        "--run-dir",
        type=Path,
        default=None,
        help="directory for handover/, snapshots/, trace.jsonl (default: trace_logs/run_<ts>)",
    )
    parser.add_argument(
        "--task-id",
        default=None,
        help="Identifier used in logs and summaries. Defaults to run_dir basename.",
    )
    parser.add_argument(
        "--empty-workbook-created",
        action="store_true",
        help="Indicates the wrapper synthesized a blank .xlsx because the task had no source workbook.",
    )
    parser.add_argument(
        "--disable-excel-mcp",
        action="store_true",
        help="Do not configure the excel-mcp server for Worker/Evaluator subprocesses.",
    )
    return parser.parse_args(argv)


def _resolve_workbook_dir(args: argparse.Namespace) -> tuple[Path | None, int]:
    if args.workbook_dir is not None:
        workbook_dir = args.workbook_dir.expanduser().resolve()
        if not workbook_dir.exists():
            print(f"workbook directory not found: {workbook_dir}", file=sys.stderr)
            return None, 2
        if not workbook_dir.is_dir():
            print(f"workbook directory is not a directory: {workbook_dir}", file=sys.stderr)
            return None, 2
        if not any(path.is_file() or path.is_symlink() for path in workbook_dir.iterdir()):
            print(f"workbook directory is empty: {workbook_dir}", file=sys.stderr)
            return None, 2
        return workbook_dir, 0

    workbooks = [path.expanduser().resolve() for path in (args.workbooks or [])]
    for workbook in workbooks:
        if not workbook.exists():
            print(f"workbook not found: {workbook}", file=sys.stderr)
            return None, 2
    try:
        return _derive_workbook_dir(workbooks), 0
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return None, 2


def _load_task(args: argparse.Namespace) -> tuple[str | None, int]:
    if args.task_json_path is None:
        return str(args.task), 0
    task_path = args.task_json_path.expanduser().resolve()
    if not task_path.is_file():
        print(f"task json not found: {task_path}", file=sys.stderr)
        return None, 2
    try:
        task_obj = json.loads(task_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"invalid JSON in {task_path}: {exc}", file=sys.stderr)
        return None, 2
    if not isinstance(task_obj, dict):
        print(f"task JSON must be an object: {task_path}", file=sys.stderr)
        return None, 2
    return json.dumps(task_obj, indent=2, ensure_ascii=False), 0


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    root = Path(__file__).resolve().parent
    run_dir = (args.run_dir or (root / "trace_logs" / f"run_{datetime.now():%Y%m%d_%H%M%S}")).resolve()

    task_text, code = _load_task(args)
    if code:
        return code
    assert task_text is not None

    workbook_dir, code = _resolve_workbook_dir(args)
    if code:
        return code
    assert workbook_dir is not None

    env_error = validate_child_env()
    if env_error is not None:
        print(env_error, file=sys.stderr)
        return 2

    cleanup_before_run(run_dir, workbook_dir=workbook_dir, remove_run_contents=False)
    run_dir.mkdir(parents=True, exist_ok=True)
    clear_run_owned_dirs(run_dir)
    unlink_or_truncate_run_file(run_dir / "trace.jsonl")
    unlink_or_truncate_run_file(run_dir / "wrapper_summary.json")

    task_id = args.task_id or run_dir.name
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
    except Exception as exc:
        summary = {
            "verdict": "fail",
            "task_id": task_id,
            "run_dir": str(run_dir),
            "final_dir": str(run_dir / "final_result"),
            "trace": str(run_dir / "trace.jsonl"),
            "error": str(exc),
        }
        (run_dir / "wrapper_summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 1
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
    (run_dir / "wrapper_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if result.verdict == "success" else 1


if __name__ == "__main__":
    raise SystemExit(main())
