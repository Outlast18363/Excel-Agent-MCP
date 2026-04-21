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
    def __init__(self, task: str, workbook: Path, run_dir: Path, task_id: str):
        self.task = task
        self.workbook = Path(workbook).resolve()
        self.run_dir = Path(run_dir).resolve()
        self.task_id = task_id

        self.handover = self.run_dir / "handover" # where handover docs are stored
        self.handover.mkdir(parents=True, exist_ok=True)
        self.plan_path = self.handover / "plan.md"
        self.impl_path = self.handover / "impl_report.md"
        self.eval_path = self.handover / "eval_report.md"
        self.hint_path = self.handover / "execution_hint.md"

        self.snapshot_path = self.run_dir / "snapshot.xlsx"
        self.trace_path = self.run_dir / "trace.jsonl"

        # final_result/ is the orchestrator-owned deliverable surface. The
        # workbook-copy name carries a `{task_id}_` signature so any copy in
        # this folder that does NOT match that prefix is provably not ours
        # (e.g. a rogue Executor copy would stand out in logs).
        self.final_dir = self.run_dir / "final_result"
        self.final_dir.mkdir(parents=True, exist_ok=True)
        self.final_workbook = self.final_dir / f"{task_id}_final_result{self.workbook.suffix}"

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

    def run(self) -> RunResult:
        with EventBus(self.trace_path) as bus:
            # Planner: produce plan.md. Workspace is the run_dir so relative paths
            # the agent writes land in a predictable place; it still gets absolute
            # workbook + handover paths via ctx.
            planner = PlannerAgent(bus, self.run_dir)
            executor = ExecutorAgent(bus, self.run_dir)
            evaluator = EvaluatorAgent(bus, self.run_dir)
            distiller = DistillerAgent(bus, self.run_dir)

            bus.transition(None, "Planner", 0, "start")
            planner.run(task=self.task, workbook=self.workbook, plan_path=self.plan_path)

            # Snapshot AFTER the plan is written but BEFORE any mutation.
            shutil.copy2(self.workbook, self.snapshot_path)
            bus.emit("ORCHESTRATOR", {"type": "snapshot", "path": str(self.snapshot_path)})

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
                    final_dir=self.final_dir,
                    impl_path=self.impl_path,
                    impl_path_or_none=self.impl_path if has_prior_impl else None,
                    eval_path_or_none=self.eval_path if has_prior_eval else None,
                    hint_path_or_none=self.hint_path if has_prior_hint else None,
                )

                # Orchestrator publishes the workbook copy unconditionally
                # *after* the Executor returns. This removes any redo-staleness
                # risk: even if the Executor only edited the workbook (no
                # standalone report), the Evaluator still sees a fresh copy.
                shutil.copy2(self.workbook, self.final_workbook)
                bus.emit("ORCHESTRATOR", {
                    "type": "final_workbook_copy",
                    "path": str(self.final_workbook),
                })

                # ---- Evaluator ----
                bus.transition("Executor", "Evaluator", iter_idx, "evaluate")
                _, verdict = evaluator.run(
                    plan_path=self.plan_path,
                    impl_path=self.impl_path,
                    workbook=self.final_workbook,
                    final_dir=self.final_dir,
                    eval_path=self.eval_path,
                    task=self.task
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

                shutil.copy2(self.snapshot_path, self.workbook)
                # Stale impl report describes operations on a workbook state that
                # no longer exists; delete it so the next Executor can't mis-read it.
                self.impl_path.unlink(missing_ok=True)
                reset_count += 1
                redo_count = 0
                reset_mode = True
                # Next loop iteration's transition should read "Distiller -> Executor".
                prev_agent = "Distiller"
                bus.reset(iter_idx, reset_count, verdict)
