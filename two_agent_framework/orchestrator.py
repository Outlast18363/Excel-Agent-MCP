"""Worker -> Evaluator orchestration with reset-only distillation."""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from .agent import DistillerAgent, EvaluatorAgent, WorkerAgent
from .event_bus import EventBus, snapshot_role_usage, usage_delta


def unlink_or_truncate_run_file(path: Path) -> None:
    """Remove *path* so a re-run starts clean.

    On Windows, ``unlink`` raises ``PermissionError`` if another process still has
    the file open (common for ``trace.jsonl``); truncate instead.
    """
    try:
        path.unlink(missing_ok=True)
    except PermissionError:
        if path.exists():
            try:
                path.open("w", encoding="utf-8").close()
            except OSError:
                pass


MAX_REDO = 2
MAX_RESET = 1
MAX_INCOMPLETE_WORKER_RETRY = 1
_ROLES = ("Worker", "Evaluator", "Distiller")


@dataclass
class RunResult:
    verdict: str
    iterations: int
    redo_count: int
    reset_count: int
    trace_path: Path
    run_dir: Path
    final_dir: Path
    usage_by_agent: dict


class Orchestrator:
    def __init__(
        self,
        task: str,
        workbooks: list[Path] | None = None,
        *,
        workbook_dir: Path | None = None,
        empty_workbook_created: bool,
        run_dir: Path,
        task_id: str,
        with_excel_mcp: bool = True,
    ):
        self.task = task
        self.with_excel_mcp = with_excel_mcp

        if workbook_dir is not None:
            self.workbook_dir = Path(workbook_dir).resolve()
            if not self.workbook_dir.is_dir():
                raise ValueError("workbook_dir must exist and be a directory.")
            self.workbooks = self._scan_workbook_dir(self.workbook_dir)
            if not self.workbooks:
                raise ValueError("workbook_dir must contain at least one staged task file.")
        else:
            resolved_workbooks = [Path(path).resolve() for path in workbooks or []]
            if not resolved_workbooks:
                raise ValueError("workbooks must contain at least one staged task file.")
            parents = {path.parent for path in resolved_workbooks}
            if len(parents) != 1:
                raise ValueError("workbooks must all live in the same staged workbook directory.")
            self.workbooks = resolved_workbooks
            self.workbook_dir = resolved_workbooks[0].parent

        self.xlsx_files = [path for path in self.workbooks if path.suffix.lower() == ".xlsx"]
        self.staging_note = self._build_staging_note(self.workbooks, empty_workbook_created)
        self.run_dir = Path(run_dir).resolve()
        self.task_id = task_id

        self.handover = self.run_dir / "handover"
        self.handover.mkdir(parents=True, exist_ok=True)
        self.impl_path = self.handover / "impl_report.md"
        self.eval_path = self.handover / "eval_report.md"
        self.hint_path = self.handover / "execution_hint.md"
        self.screenshots_dir = self.run_dir / "screenshots"

        self.snapshot_dir = self.run_dir / "snapshots"
        self.snapshot_dir.mkdir(parents=True, exist_ok=True)
        self.trace_path = self.run_dir / "trace.jsonl"

        self.final_dir = self.run_dir / "final_result"
        self.final_dir.mkdir(parents=True, exist_ok=True)

        self._clear_prior_run()

    @staticmethod
    def _scan_workbook_dir(workbook_dir: Path) -> list[Path]:
        return sorted(
            [path.resolve() for path in workbook_dir.iterdir() if path.is_file() or path.is_symlink()],
            key=lambda path: path.name.lower(),
        )

    @staticmethod
    def _build_staging_note(workbooks: list[Path], empty_workbook_created: bool) -> str:
        if empty_workbook_created:
            return (
                "Workbook note: the primary .xlsx was created by the orchestrator as a blank "
                "workbook because the task did not provide one."
            )
        suffixes = sorted({path.suffix.lower() or "(no extension)" for path in workbooks})
        if not suffixes:
            return "Staging note: the workbook folder contains the staged task files."
        return f"Staging note: staged task files include {', '.join(suffixes)}."

    @staticmethod
    def _build_workbook_note(empty_workbook_created: bool) -> str:
        return Orchestrator._build_staging_note([], empty_workbook_created)

    @staticmethod
    def _usage_with_roles(usage_by_agent: dict[str, dict[str, int]]) -> dict[str, dict]:
        return {role: dict(usage_by_agent.get(role, {})) for role in _ROLES}

    def _clear_prior_run(self) -> None:
        old_single_snapshot = self.run_dir / "snapshot.xlsx"
        owned_dirs = [self.handover, self.final_dir, self.snapshot_dir, self.screenshots_dir]
        has_prior_artifacts = old_single_snapshot.exists() or any(
            folder.exists() and any(folder.iterdir()) for folder in owned_dirs
        )
        if not has_prior_artifacts:
            return
        unlink_or_truncate_run_file(self.trace_path)
        for folder in owned_dirs:
            if not folder.exists():
                continue
            for path in folder.iterdir():
                if path.is_file() or path.is_symlink():
                    path.unlink()
                else:
                    shutil.rmtree(path)
        staged_paths = {(self.workbook_dir / source_path.name).resolve() for source_path in self.workbooks}
        if self.workbook_dir != self.run_dir:
            for path in self.workbook_dir.iterdir():
                if path.resolve() in staged_paths:
                    continue
                if path.is_file() or path.is_symlink():
                    path.unlink()
                else:
                    shutil.rmtree(path)
        old_single_snapshot.unlink(missing_ok=True)

    def _wipe_final_dir(self) -> None:
        for path in self.final_dir.iterdir():
            if path.is_file() or path.is_symlink():
                path.unlink()
            else:
                shutil.rmtree(path)

    def _publish_workbook_files(self) -> list[Path]:
        published: list[Path] = []
        if self.workbook_dir != self.run_dir and self.workbook_dir.is_dir():
            for path in sorted(self.workbook_dir.iterdir(), key=lambda p: p.name.lower()):
                if not (path.is_file() or path.is_symlink()):
                    continue
                target = self.final_dir / path.name
                shutil.copy2(path, target)
                published.append(target)
            return published
        for workbook_path in self.workbooks:
            target = self.final_dir / workbook_path.name
            shutil.copy2(workbook_path, target)
            published.append(target)
        return published

    def _snapshot_staged_files(self) -> list[Path]:
        snapshot_files: list[Path] = []
        for workbook_path in self.workbooks:
            snapshot_path = self.snapshot_dir / workbook_path.name
            shutil.copy2(workbook_path, snapshot_path)
            snapshot_files.append(snapshot_path)
        return snapshot_files

    def _restore_workbook_dir_from_snapshots(self) -> None:
        snapshot_files = [path for path in self.snapshot_dir.iterdir() if path.is_file() or path.is_symlink()]
        if self.workbook_dir == self.run_dir:
            for snapshot in snapshot_files:
                shutil.copy2(snapshot, self.workbook_dir / snapshot.name)
            return
        for path in self.workbook_dir.iterdir():
            if path.is_file() or path.is_symlink():
                path.unlink()
            else:
                shutil.rmtree(path)
        for snapshot in snapshot_files:
            shutil.copy2(snapshot, self.workbook_dir / snapshot.name)

    def _result(
        self,
        verdict: str,
        iterations: int,
        redo_count: int,
        reset_count: int,
        bus: EventBus,
    ) -> RunResult:
        return RunResult(
            verdict,
            iterations,
            redo_count,
            reset_count,
            self.trace_path,
            self.run_dir,
            self.final_dir,
            self._usage_with_roles(bus.usage_by_agent),
        )

    def _worker_completion_issues(self) -> list[str]:
        if not self.impl_path.exists():
            return [f"missing implementation report: {self.impl_path}"]
        try:
            report = self.impl_path.read_text(encoding="utf-8")
        except OSError as exc:
            return [f"could not read implementation report: {exc}"]
        issues = []
        for heading in ("# Execution Todo", "# Implementation Report"):
            if heading not in report:
                issues.append(f"implementation report missing {heading!r}")
        return issues

    def run(self) -> RunResult:
        with EventBus(self.trace_path) as bus:
            worker = WorkerAgent(bus, self.run_dir, with_excel_mcp=self.with_excel_mcp)
            evaluator = EvaluatorAgent(bus, self.run_dir, with_excel_mcp=self.with_excel_mcp)
            distiller = DistillerAgent(bus, self.run_dir, with_excel_mcp=self.with_excel_mcp)

            snapshot_files = self._snapshot_staged_files()
            bus.snapshot(self.snapshot_dir, snapshot_files, self.workbooks)

            redo_count = 0
            reset_count = 0
            iter_idx = 0
            prev_agent: str | None = None
            reset_mode = False

            while True:
                self._wipe_final_dir()
                bus.final_output_dir_wipe(iter_idx, self.final_dir)

                reason = "reset-retry" if reset_mode else ("redo" if iter_idx > 0 else "start")
                bus.transition(prev_agent, "Worker", iter_idx, reason)
                u0 = snapshot_role_usage("Worker", bus.usage_by_agent)
                worker_issues: list[str] = []
                for attempt in range(MAX_INCOMPLETE_WORKER_RETRY + 1):
                    task = self.task
                    if attempt:
                        task = (
                            f"{self.task}\n\n"
                            "Previous Worker invocation ended before completing required side effects: "
                            f"{'; '.join(worker_issues)}. Continue now from the staged workbook. "
                            "Write the implementation report and complete the workbook task before your final response."
                        )
                    worker.run(
                        task=task,
                        workbook_dir=self.workbook_dir,
                        staging_note=self.staging_note,
                        final_dir=self.final_dir,
                        impl_path=self.impl_path,
                        eval_path_or_none=None if reset_mode or iter_idx == 0 else self.eval_path,
                        hint_path_or_none=self.hint_path if reset_mode else None,
                        run_dir=self.run_dir,
                    )
                    worker_issues = self._worker_completion_issues()
                    if not worker_issues:
                        break
                    bus.emit("ORCHESTRATOR", {
                        "type": "worker_incomplete",
                        "iter": iter_idx,
                        "attempt": attempt,
                        "issues": worker_issues,
                    })
                worker_usage = usage_delta(snapshot_role_usage("Worker", bus.usage_by_agent), u0)
                if worker_issues:
                    bus.agent_session_completed("Worker", iter_idx, worker_usage)
                    return self._result("fail", iter_idx, redo_count, reset_count, bus)

                published_files = self._publish_workbook_files()
                bus.final_workbook_files_copy(iter_idx, self.final_dir, published_files)

                bus.agent_session_completed("Worker", iter_idx, worker_usage)
                bus.transition("Worker", "Evaluator", iter_idx, "evaluate")
                u0 = snapshot_role_usage("Evaluator", bus.usage_by_agent)
                _, verdict = evaluator.run(
                    task=self.task,
                    snapshot_dir=self.snapshot_dir,
                    final_dir=self.final_dir,
                    impl_path=self.impl_path,
                    eval_path=self.eval_path,
                    staging_note=self.staging_note,
                    run_dir=self.run_dir,
                )
                eval_usage = usage_delta(snapshot_role_usage("Evaluator", bus.usage_by_agent), u0)
                bus.agent_session_completed("Evaluator", iter_idx, eval_usage)
                bus.verdict(verdict, iter_idx)

                iter_idx += 1
                bus.iteration_usage(iter_idx, self._usage_with_roles(bus.usage_by_agent))
                prev_agent = "Evaluator"
                reset_mode = False

                if verdict == "success":
                    return self._result("success", iter_idx, redo_count, reset_count, bus)

                if verdict == "redo" and redo_count < MAX_REDO:
                    redo_count += 1
                    continue

                if reset_count >= MAX_RESET:
                    return self._result("fail", iter_idx, redo_count, reset_count, bus)

                bus.transition("Evaluator", "Distiller", iter_idx, "distill")
                bus.distill(iter_idx, self.eval_path, self.hint_path)
                u0 = snapshot_role_usage("Distiller", bus.usage_by_agent)
                distiller.run(eval_path=self.eval_path, hint_path=self.hint_path)
                dist_usage = usage_delta(snapshot_role_usage("Distiller", bus.usage_by_agent), u0)
                bus.agent_session_completed("Distiller", iter_idx, dist_usage)

                self._restore_workbook_dir_from_snapshots()
                bus.restore_from_snapshot(iter_idx, self.snapshot_dir, self.workbook_dir)
                self.impl_path.unlink(missing_ok=True)
                reset_count += 1
                redo_count = 0
                reset_mode = True
                prev_agent = "Distiller"
                bus.reset(iter_idx, reset_count, verdict)
