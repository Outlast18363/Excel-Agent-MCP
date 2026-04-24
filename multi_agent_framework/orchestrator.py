"""Planner → Executor ⇄ Evaluator loop.

Loop behavior is tuned by two module-level constants. The Planner runs once,
then we snapshot the workbook, then we alternate Executor/Evaluator. A `redo`
verdict re-runs the Executor in place; a `reset` (or a redo cap hit) rolls the
workbook back to the snapshot and hands the Evaluator's report to a fresh
Executor attempt.
"""

import shutil
from dataclasses import dataclass
from pathlib import Path

from .agent import DistillerAgent, EvaluatorAgent, ExecutorAgent, PlannerAgent
from .event_bus import EventBus

MAX_REDO = 3   # successive in-place fix attempts before we're forced to reset
MAX_RESET = 1  # how many times we may roll the workbook back and restart


@dataclass
class RunResult:
    verdict: str          # "success" | "fail"
    iterations: int       # total Executor/Evaluator rounds completed
    redo_count: int
    reset_count: int
    trace_path: Path
    run_dir: Path
    final_dir: Path
    # Per-agent cumulative token usage (collected from turn.completed events).
    usage_by_agent: dict = None  # type: ignore[assignment]


class Orchestrator:
    def __init__(
        self,
        task: str,
        workbooks: list[Path],
        empty_workbook_created: bool,
        run_dir: Path,
        task_id: str,
    ):
        self.task = task
        self.workbooks = [Path(path).resolve() for path in workbooks]
        if not self.workbooks:
            raise ValueError("workbooks must contain at least one staged task file.")
        self.workbook = next(
            (path for path in self.workbooks if path.suffix.lower() == ".xlsx"),
            None,
        )
        if self.workbook is None:
            raise ValueError("workbooks must include at least one .xlsx file.")
        self.workbook_dir = self.workbook.parent
        self.workbook_note = self._build_workbook_note(empty_workbook_created)
        self.run_dir = Path(run_dir).resolve()
        self.task_id = task_id

        self.handover = self.run_dir / "handover" # where handover docs are stored
        self.handover.mkdir(parents=True, exist_ok=True)
        self.plan_path = self.handover / "plan.md"
        self.impl_path = self.handover / "impl_report.md"
        self.eval_path = self.handover / "eval_report.md"
        self.hint_path = self.handover / "execution_hint.md"
        self.screenshots_dir = self.run_dir / "screenshots"

        self.snapshot_dir = self.run_dir / "snapshots"
        self.snapshot_dir.mkdir(parents=True, exist_ok=True)
        self.snapshot_path = self.snapshot_dir / self.workbook.name
        self.trace_path = self.run_dir / "trace.jsonl"

        # final_result/ is the orchestrator-owned deliverable surface.
        self.final_dir = self.run_dir / "final_result"
        self.final_dir.mkdir(parents=True, exist_ok=True)
        self.final_workbook = self.final_dir / self.workbook.name

        # If this run_dir has been used before, wipe the prior run's artifacts so
        # the new run starts from a clean slate. Must run AFTER the mkdir calls
        # above so the target folders are guaranteed to exist when we iterate.
        self._clear_prior_run()

    @staticmethod
    def _build_workbook_note(empty_workbook_created: bool) -> str:
        """Describe whether the primary workbook came from the task or was synthesized."""
        if empty_workbook_created:
            return (
                "Workbook note: the primary .xlsx was created by the orchestrator as a blank "
                "workbook because the task did not provide one."
            )
        return "Workbook note: the primary .xlsx came from the staged task files."

    def _clear_prior_run(self) -> None:
        """Wipe artifacts from a prior orchestrator run in this run_dir, if any.

        Detection: non-empty orchestrator-owned folders from a prior run
        (`snapshots/`, `handover/`, `final_result/`, or `screenshots/`) or the
        legacy top-level `snapshot.xlsx` file identify a reused run_dir. When
        detected, clear the run-owned folders fully, clear the workbook/
        folder except for the caller-staged source files, and delete stale
        snapshots so run() can copy fresh ones.
        """
        old_single_snapshot = self.run_dir / "snapshot.xlsx"
        owned_dirs = [self.handover, self.final_dir, self.snapshot_dir, self.screenshots_dir]
        has_prior_artifacts = old_single_snapshot.exists() or any(
            folder.exists() and any(folder.iterdir()) for folder in owned_dirs
        )
        if not has_prior_artifacts:
            return
        # Folders fully owned by the orchestrator — safe to wipe all contents.
        for folder in owned_dirs:
            if not folder.exists():
                continue
            for p in folder.iterdir():
                if p.is_file() or p.is_symlink():
                    p.unlink()
                else:
                    shutil.rmtree(p)
        # Workbook folder is caller-staged: preserve the current task files but
        # remove anything else lingering from a prior task. Guarded against the
        # degenerate case where workbook lives directly in run_dir.
        staged_paths = {
            (self.workbook_dir / source_path.name).resolve()
            for source_path in self.workbooks
        }
        if self.workbook_dir != self.run_dir:
            for p in self.workbook_dir.iterdir():
                if p.resolve() in staged_paths:
                    continue
                if p.is_file() or p.is_symlink():
                    p.unlink()
                else:
                    shutil.rmtree(p)
        old_single_snapshot.unlink(missing_ok=True)

    def _wipe_final_dir(self) -> None:
        """Remove everything inside final_dir without deleting the folder itself.

        Called at the top of every loop iteration so stale non-xlsx reports
        (from a prior Executor run) and the prior workbook copy never leak
        into the next Evaluator pass. Idempotent when the folder is empty.
        """
        for p in self.final_dir.iterdir():
            if p.is_file() or p.is_symlink():
                p.unlink()
            else:
                shutil.rmtree(p)

    def _publish_workbook_files(self) -> None:
        """Copy staged workbook-folder files into final_result/ without deleting other outputs."""
        if self.workbook_dir != self.run_dir and self.workbook_dir.is_dir():
            for p in self.workbook_dir.iterdir():
                if not (p.is_file() or p.is_symlink()):
                    continue
                shutil.copy2(p, self.final_dir / p.name)
            return
        # Degenerate standalone fallback: only copy the tracked staged files so
        # we do not mirror unrelated run_dir artifacts into final_result/.
        for workbook_path in self.workbooks:
            shutil.copy2(workbook_path, self.final_dir / workbook_path.name)

    def _restore_workbook_dir_from_snapshots(self) -> None:
        """Restore every staged source file from snapshots/ into workbook/."""
        snapshot_files = [p for p in self.snapshot_dir.iterdir() if p.is_file() or p.is_symlink()]
        if self.workbook_dir == self.run_dir:
            # Degenerate standalone mode: never wipe run_dir itself because it
            # also contains handover/, final_result/, snapshots/, and trace.jsonl.
            for snapshot in snapshot_files:
                shutil.copy2(snapshot, self.workbook_dir / snapshot.name)
            return
        for p in self.workbook_dir.iterdir():
            if p.is_file() or p.is_symlink():
                p.unlink()
            else:
                shutil.rmtree(p)
        for snapshot in snapshot_files:
            shutil.copy2(snapshot, self.workbook_dir / snapshot.name)

    def run(self) -> RunResult:
        with EventBus(self.trace_path) as bus:
            # Planner: produce plan.md. Workspace is the run_dir so relative paths
            # the agent writes land in a predictable place; it still gets absolute
            # workbook + handover paths via ctx.
            planner = PlannerAgent(bus, self.run_dir)
            executor = ExecutorAgent(bus, self.run_dir)
            evaluator = EvaluatorAgent(bus, self.run_dir)
            distiller = DistillerAgent(bus, self.run_dir)

            # Snapshot BEFORE the Planner runs, from the staged task files in
            # workbook/. This guarantees the rollback targets are untouched by
            # any agent, including the Planner whose xlwings-backed inspection
            # tools could otherwise affect file state on disk.
            snapshot_files: list[str] = []
            for workbook_path in self.workbooks:
                snapshot_path = self.snapshot_dir / workbook_path.name
                shutil.copy2(workbook_path, snapshot_path)
                snapshot_files.append(str(snapshot_path))
            bus.emit("ORCHESTRATOR", {
                "type": "snapshot",
                "path": str(self.snapshot_path),
                "paths": snapshot_files,
                "workbooks": [str(p) for p in self.workbooks],
            })

            bus.transition(None, "Planner", 0, "start")
            planner.run(
                task=self.task,
                workbook=self.workbook,
                workbook_dir=self.workbook_dir,
                workbook_note=self.workbook_note,
                plan_path=self.plan_path,
                run_dir=self.run_dir,
            )

            redo_count = 0
            reset_count = 0
            reset_mode = False     # True on the iteration immediately following a reset
            iter_idx = 0
            prev_agent = "Planner"

            while True:
                # Wipe final_result/ at the top of every iteration (initial,
                # redo, reset). Single invariant: the Evaluator only ever sees
                # artifacts produced by *this* iteration.
                self._wipe_final_dir()
                bus.emit("ORCHESTRATOR", {"type": "final_output_dir_wipe", "iter": iter_idx})

                # ---- Executor ----
                reason = "reset-retry" if reset_mode else ("redo" if iter_idx > 0 else "execute")
                bus.transition(prev_agent, "Executor", iter_idx, reason)
                # On the very first iteration there's no prior impl/eval/hint.
                # On a redo, the Executor sees the prior impl + eval reports.
                # On a reset-retry, impl_report.md has been deleted and the full
                # eval report is replaced by the distilled execution_hint.md.
                has_prior_impl = iter_idx > 0 and not reset_mode
                has_prior_eval = iter_idx > 0 and not reset_mode
                has_prior_hint = reset_mode
                executor.run(
                    task=self.task,
                    plan_path=self.plan_path,
                    workbook=self.workbook,
                    workbook_dir=self.workbook_dir,
                    workbook_note=self.workbook_note,
                    final_dir=self.final_dir,
                    impl_path=self.impl_path,
                    impl_path_or_none=self.impl_path if has_prior_impl else None,
                    eval_path_or_none=self.eval_path if has_prior_eval else None,
                    hint_path_or_none=self.hint_path if has_prior_hint else None,
                    run_dir=self.run_dir,
                )

                # Orchestrator publishes the staged workbook-folder files
                # *after* the Executor returns. This keeps the Evaluator aligned
                # with the current workbook state without deleting any
                # non-Excel deliverables the Executor already wrote.
                self._publish_workbook_files()
                bus.emit("ORCHESTRATOR", {
                    "type": "final_workbook_files_copy",
                    "path": str(self.final_dir),
                    "files": [str(self.final_dir / p.name) for p in self.workbooks],
                })

                # ---- Evaluator ----
                bus.transition("Executor", "Evaluator", iter_idx, "evaluate")
                _, verdict = evaluator.run(
                    plan_path=self.plan_path,
                    impl_path=self.impl_path,
                    workbook=self.final_workbook,
                    workbook_dir=self.workbook_dir,
                    workbook_note=self.workbook_note,
                    final_dir=self.final_dir,
                    eval_path=self.eval_path,
                    task=self.task,
                    run_dir=self.run_dir,
                )
                bus.verdict(verdict, iter_idx)

                iter_idx += 1
                prev_agent = "Evaluator"

                if verdict == "success":
                    return RunResult(
                        "success", iter_idx, redo_count, reset_count,
                        self.trace_path, self.run_dir, self.final_dir,
                        dict(bus.usage_by_agent),
                    )

                # Decide redo vs reset. A `redo` verdict only escalates to reset
                # once the redo cap is exhausted; a `reset` verdict goes straight there.
                if verdict == "redo" and redo_count < MAX_REDO:
                    redo_count += 1
                    reset_mode = False
                    continue

                # Reset path (verdict=="reset", or verdict=="redo" with cap hit).
                if reset_count >= MAX_RESET:
                    return RunResult(
                        "fail", iter_idx, redo_count, reset_count,
                        self.trace_path, self.run_dir, self.final_dir,
                        dict(bus.usage_by_agent),
                    )

                # Distill BEFORE rollback: the Distiller reads eval_report.md while
                # it still describes the about-to-be-reverted state. Rollback and
                # impl_report.md deletion happen after, so the only reset-surviving
                # handover docs are plan.md and the freshly-written execution_hint.md.
                bus.transition("Evaluator", "Distiller", iter_idx, "distill")
                distiller.run(eval_path=self.eval_path, hint_path=self.hint_path)

                self._restore_workbook_dir_from_snapshots()
                # Stale impl report describes operations on a workbook state that
                # no longer exists; delete it so the next Executor can't mis-read it.
                self.impl_path.unlink(missing_ok=True)
                reset_count += 1
                redo_count = 0
                reset_mode = True
                # Next loop iteration's transition should read "Distiller -> Executor".
                prev_agent = "Distiller"
                bus.reset(iter_idx, reset_count, verdict)
