from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from pathlib import Path

from claude_invocation import run_claude_role
from trace import EventBus, normalize_usage, snapshot_role_usage, usage_delta


MAX_REDO = 2
MAX_RESET = 1
MAX_INCOMPLETE_WORKER_RETRY = 1
ROLES = ("Worker", "Evaluator", "Distiller")


def unlink_or_truncate_run_file(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except PermissionError:
        if path.exists():
            try:
                path.open("w", encoding="utf-8").close()
            except OSError:
                pass


@dataclass
class RunResult:
    verdict: str
    iterations: int
    redo_count: int
    reset_count: int
    trace_path: Path
    run_dir: Path
    final_dir: Path
    usage_by_agent: dict[str, dict[str, int]]


class WorkerAgent:
    def __init__(self, bus: EventBus, run_dir: Path, *, with_excel_mcp: bool) -> None:
        self.bus = bus
        self.run_dir = run_dir
        self.with_excel_mcp = with_excel_mcp

    def run(
        self,
        task: str,
        workbook_dir: Path,
        staging_note: str,
        final_dir: Path,
        impl_path: Path,
        eval_path_or_none: Path | None,
        hint_path_or_none: Path | None,
        run_dir: Path,
    ) -> str:
        prompt = build_worker_prompt(
            task=task,
            workbook_dir=workbook_dir,
            staging_note=staging_note,
            final_dir=final_dir,
            impl_path=impl_path,
            eval_path_or_none=None if hint_path_or_none is not None else eval_path_or_none,
            hint_path_or_none=hint_path_or_none,
            run_dir=run_dir,
        )
        return run_claude_role(
            "Worker",
            prompt,
            bus=self.bus,
            run_dir=self.run_dir,
            workbook_dir=workbook_dir,
            with_excel_mcp=self.with_excel_mcp,
        ).final_text


class EvaluatorAgent:
    VERDICT_RE = re.compile(r"<verdict>\s*(success|redo|reset)\s*</verdict>", re.IGNORECASE)
    MAX_VERDICT_RETRY = 2

    def __init__(self, bus: EventBus, run_dir: Path, *, with_excel_mcp: bool) -> None:
        self.bus = bus
        self.run_dir = run_dir
        self.with_excel_mcp = with_excel_mcp

    def run(
        self,
        task: str,
        snapshot_dir: Path,
        final_dir: Path,
        impl_path: Path,
        eval_path: Path,
        staging_note: str,
        run_dir: Path,
    ) -> tuple[str, str]:
        prompt = build_evaluator_prompt(
            task=task,
            snapshot_dir=snapshot_dir,
            final_dir=final_dir,
            impl_path=impl_path,
            eval_path=eval_path,
            staging_note=staging_note,
            run_dir=run_dir,
        )
        final_msg = ""
        for attempt in range(self.MAX_VERDICT_RETRY + 1):
            if attempt:
                prompt = (
                    f"{prompt}\n\nYour previous final message did not contain exactly one required verdict tag. "
                    "Re-read the evidence, update the evaluation report if needed, and end with exactly one verdict tag."
                )
            final_msg = run_claude_role(
                "Evaluator",
                prompt,
                bus=self.bus,
                run_dir=self.run_dir,
                workbook_dir=final_dir,
                with_excel_mcp=self.with_excel_mcp,
            ).final_text
            match = self.VERDICT_RE.search(final_msg)
            if match:
                return final_msg, match.group(1).lower()
        self.bus.emit("Evaluator", {
            "type": "warning",
            "reason": f"verdict tag missing after {self.MAX_VERDICT_RETRY + 1} attempts",
            "final_msg_tail": final_msg[-200:],
        })
        return final_msg, "redo"


class DistillerAgent:
    def __init__(self, bus: EventBus, run_dir: Path, *, with_excel_mcp: bool) -> None:
        self.bus = bus
        self.run_dir = run_dir
        self.with_excel_mcp = with_excel_mcp

    def run(self, eval_path: Path, hint_path: Path, workbook_dir: Path) -> str:
        prompt = build_distiller_prompt(eval_path=eval_path, hint_path=hint_path)
        return run_claude_role(
            "Distiller",
            prompt,
            bus=self.bus,
            run_dir=self.run_dir,
            workbook_dir=workbook_dir,
            with_excel_mcp=False,
        ).final_text


class Orchestrator:
    def __init__(
        self,
        task: str,
        *,
        workbook_dir: Path,
        empty_workbook_created: bool,
        run_dir: Path,
        task_id: str,
        with_excel_mcp: bool = True,
    ) -> None:
        self.task = task
        self.with_excel_mcp = with_excel_mcp
        self.workbook_dir = Path(workbook_dir).resolve()
        if not self.workbook_dir.is_dir():
            raise ValueError("workbook_dir must exist and be a directory.")
        self.workbooks = self._scan_workbook_dir(self.workbook_dir)
        if not self.workbooks:
            raise ValueError("workbook_dir must contain at least one staged task file.")

        self.staging_note = self._build_staging_note(self.workbooks, empty_workbook_created)
        self.run_dir = Path(run_dir).resolve()
        self.task_id = task_id

        self.handover = self.run_dir / "handover"
        self.impl_path = self.handover / "impl_report.md"
        self.eval_path = self.handover / "eval_report.md"
        self.hint_path = self.handover / "execution_hint.md"
        self.snapshot_dir = self.run_dir / "snapshots"
        self.screenshots_dir = self.run_dir / "screenshots"
        self.final_dir = self.run_dir / "final_result"
        self.trace_path = self.run_dir / "trace.jsonl"

        for path in (self.handover, self.snapshot_dir, self.screenshots_dir, self.final_dir):
            path.mkdir(parents=True, exist_ok=True)
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
        return f"Staging note: staged task files include {', '.join(suffixes)}."

    @staticmethod
    def _usage_with_roles(usage_by_agent: dict[str, dict[str, int]]) -> dict[str, dict[str, int]]:
        return {role: normalize_usage(usage_by_agent.get(role) or {}) for role in ROLES}

    def _clear_prior_run(self) -> None:
        unlink_or_truncate_run_file(self.trace_path)
        for folder in (self.handover, self.final_dir, self.snapshot_dir, self.screenshots_dir):
            if not folder.exists():
                continue
            for path in folder.iterdir():
                if path.is_file() or path.is_symlink():
                    path.unlink()
                else:
                    shutil.rmtree(path)

    def _wipe_final_dir(self) -> None:
        for path in self.final_dir.iterdir():
            if path.is_file() or path.is_symlink():
                path.unlink()
            else:
                shutil.rmtree(path)

    def _snapshot_staged_files(self) -> list[Path]:
        snapshot_files: list[Path] = []
        for source in self.workbooks:
            target = self.snapshot_dir / source.name
            shutil.copy2(source, target)
            snapshot_files.append(target)
        return snapshot_files

    def _restore_workbook_dir_from_snapshots(self) -> None:
        for path in self.workbook_dir.iterdir():
            if path.is_file() or path.is_symlink():
                path.unlink()
            else:
                shutil.rmtree(path)
        for snapshot in sorted(self.snapshot_dir.iterdir(), key=lambda p: p.name.lower()):
            if snapshot.is_file() or snapshot.is_symlink():
                shutil.copy2(snapshot, self.workbook_dir / snapshot.name)
        self.workbooks = self._scan_workbook_dir(self.workbook_dir)

    def _publish_workbook_files(self) -> list[Path]:
        published: list[Path] = []
        for path in sorted(self.workbook_dir.iterdir(), key=lambda p: p.name.lower()):
            if not (path.is_file() or path.is_symlink()):
                continue
            target = self.final_dir / path.name
            shutil.copy2(path, target)
            published.append(target)
        return published

    def _worker_completion_issues(self) -> list[str]:
        if not self.impl_path.exists():
            return [f"missing implementation report: {self.impl_path}"]
        try:
            report = self.impl_path.read_text(encoding="utf-8")
        except OSError as exc:
            return [f"could not read implementation report: {exc}"]
        issues: list[str] = []
        for heading in ("# Execution Todo", "# Implementation Report"):
            if heading not in report:
                issues.append(f"implementation report missing {heading!r}")
        return issues

    def _result(self, verdict: str, iterations: int, redo_count: int, reset_count: int, bus: EventBus) -> RunResult:
        return RunResult(
            verdict=verdict,
            iterations=iterations,
            redo_count=redo_count,
            reset_count=reset_count,
            trace_path=self.trace_path,
            run_dir=self.run_dir,
            final_dir=self.final_dir,
            usage_by_agent=self._usage_with_roles(bus.usage_by_agent),
        )

    def run(self) -> RunResult:
        with EventBus(self.trace_path) as bus:
            worker = WorkerAgent(bus, self.run_dir, with_excel_mcp=self.with_excel_mcp)
            evaluator = EvaluatorAgent(bus, self.run_dir, with_excel_mcp=self.with_excel_mcp)
            distiller = DistillerAgent(bus, self.run_dir, with_excel_mcp=False)

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
                before = snapshot_role_usage("Worker", bus.usage_by_agent)
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
                worker_usage = usage_delta(snapshot_role_usage("Worker", bus.usage_by_agent), before)
                bus.agent_session_completed("Worker", iter_idx, worker_usage)
                if worker_issues:
                    return self._result("fail", iter_idx, redo_count, reset_count, bus)

                published_files = self._publish_workbook_files()
                bus.final_workbook_files_copy(iter_idx, self.final_dir, published_files)

                bus.transition("Worker", "Evaluator", iter_idx, "evaluate")
                before = snapshot_role_usage("Evaluator", bus.usage_by_agent)
                _, verdict = evaluator.run(
                    task=self.task,
                    snapshot_dir=self.snapshot_dir,
                    final_dir=self.final_dir,
                    impl_path=self.impl_path,
                    eval_path=self.eval_path,
                    staging_note=self.staging_note,
                    run_dir=self.run_dir,
                )
                eval_usage = usage_delta(snapshot_role_usage("Evaluator", bus.usage_by_agent), before)
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
                before = snapshot_role_usage("Distiller", bus.usage_by_agent)
                distiller.run(eval_path=self.eval_path, hint_path=self.hint_path, workbook_dir=self.workbook_dir)
                dist_usage = usage_delta(snapshot_role_usage("Distiller", bus.usage_by_agent), before)
                bus.agent_session_completed("Distiller", iter_idx, dist_usage)

                self._restore_workbook_dir_from_snapshots()
                bus.restore_from_snapshot(iter_idx, self.snapshot_dir, self.workbook_dir)
                self.impl_path.unlink(missing_ok=True)
                reset_count += 1
                redo_count = 0
                reset_mode = True
                prev_agent = "Distiller"
                bus.reset(iter_idx, reset_count, verdict)


def build_worker_prompt(
    *,
    task: str,
    workbook_dir: Path,
    staging_note: str,
    final_dir: Path,
    impl_path: Path,
    eval_path_or_none: Path | None,
    hint_path_or_none: Path | None,
    run_dir: Path,
) -> str:
    return f"""\
You are the WORKER. You have to implement the spreadsheet task, and you cannot end before you have completed the task.

## Required workflow
1. Before substantial edits, write an `# Execution Todo` section to a markdown file at {impl_path}.
2. Todo items must be actionable high-level implementation steps, not an exhaustive data inventory.
3. Mark each completed todo item with `[CHECKED]` as you complete it.
4. After implementation, complete the `# Implementation Report` section in {impl_path}.
5. Do not copy workbook/source files into {final_dir}; the orchestrator publishes staged files there after you finish.
6. If the task requires a non-xlsx deliverable, write that deliverable directly into {final_dir}.
7. ALWAYS close any workbook you open when you are done.
8. You HAVE to explicitly calculate the number of cells in the range you are able to query (such as calculating that A1:B100 has 200 cells). If the cell count is greater than 100, you HAVE to use the `mcp__excel__local_screenshot` tool, or summarize through a narrow `mcp__excel__get_range` query of only the specific data needed.

## Required report format
# Execution Todo

- [ ] Step 1: ...
- [ ] Step 2: ...
- [ ] Step 3: ...

# Implementation Report

## Completed Work
...

## Changed Outputs
...

## Verification Performed
...

## Known Ambiguities
...

## Tool guidance
- Use `mcp__excel__local_screenshot` to quickly understand workbook layout or for range checks with > 100 cells; store screenshots under {run_dir}/screenshots/.
- Use Python/openpyxl for file edits and small range inspection, and use xlwings or Excel MCP when workbook rendering, charts, or calculated Excel behavior matter.
- Claude Code shell access is through PowerShell, not Bash. Use PowerShell-compatible commands.

## Redo/reset guidance
- Prior evaluation report for redo: {eval_path_or_none}
- Distilled reset hint: {hint_path_or_none}
- On redo, read the full evaluation report and fix only evaluator-identified issues unless a direct dependency requires more.
- On reset, use only the distilled hint above; do not rely on stale implementation or evaluation history.

## Important information
Task record (JSON):
{task}
Workbook/source folder to modify: {workbook_dir}
{staging_note}
Final deliverable folder: {final_dir}
Implementation report path: {impl_path}
Run directory: {run_dir}
"""


def build_evaluator_prompt(
    *,
    task: str,
    snapshot_dir: Path,
    final_dir: Path,
    impl_path: Path,
    eval_path: Path,
    staging_note: str,
    run_dir: Path,
) -> str:
    return f"""\
You are the EVALUATOR. Verify the Worker's output against the original task.

You may use only:
- Original files in the snapshots folder: {snapshot_dir}
- Final outputs in the final_result folder: {final_dir}
- Task record (JSON): {task}
- Implementation report: {impl_path}

Staging note:
{staging_note}

Write your evaluation report to: {eval_path}

## Required report format
# Evaluation Summary

# Checks

# Issues

# Verdict

## Verdict rules
- `success`: task is satisfied. Minor harmless issues should not trigger redo.
- `redo`: localized defects can be fixed in the current workbook.
- `reset`: workbook state is broadly corrupted, interdependent mistakes exist, or in-place repair is riskier than restoring snapshot.

## Behavior requirements
1. If non-xlsx files exist in {final_dir}, inspect them and treat them as primary deliverables when present.
2. Compare final outputs to original files in {snapshot_dir}, never to the mutable staging folder.
3. Use the implementation report at {impl_path} to understand what the Worker claims changed.
4. When using `mcp__excel__local_screenshot`, store screenshots under {run_dir}/screenshots/.
5. ALWAYS close any workbook you open when you are done.
6. You may use the reference file to evaluate the output of the Worker.

In your FINAL assistant message, end with exactly one verdict tag on its own:
<verdict>success</verdict>
<verdict>redo</verdict>
<verdict>reset</verdict>
"""


def build_distiller_prompt(eval_path: Path, hint_path: Path) -> str:
    eval_report = eval_path.read_text(encoding="utf-8") if eval_path.exists() else ""
    return f"""\
## ROLE
You are the **Distiller Agent**. You act as the bridge between an Evaluator (who audits tasks) and a Worker (who performs tasks). You will receive the Evaluator's "Verbose Evaluation Report" below. You must process this report and generate a standalone file to guide the *next* Worker.

## OBJECTIVE
Extract the core methodological failures from the verbose evaluation report and synthesize them into strict technical guardrails. The goal is to prevent the new Worker from repeating the same logical, technical, or syntactical mistakes, without cluttering their context with irrelevant data from a file state that no longer exists.

## DISTILLATION RULES (CRITICAL)
1. **Assume a Clean Slate (No Ghost States):** The new Worker is working on a fresh, reset file. Do NOT reference specific corrupted values, specific cell errors (e.g., "Cell A15 was #REF!"), or historical artifacts from the failed run.
2. **Abstract to Methodology:** Shift from *what* went wrong to *why* it went wrong.
    * *Wrong:* "The worker put the wrong tax rate in Row 10."
    * *Right:* "Failure to use absolute referencing for static variables across a range."
3. **Strict Binaries:** You must categorize all insights into two sections: **Fatal Pitfalls** (what to avoid) and **Required Mitigation Strategies** (the exact technical solution).
4. **Tool & Syntax Specificity:** Provide explicit guidance on `xlwings` usage, Python logic, or Excel formula syntax if the Evaluator identified tool-related failures.

<verbose_evaluation_report>
{eval_report}
</verbose_evaluation_report>
---

## OUTPUT TEMPLATE
You must output your response strictly in the following format. Do not include introductory or concluding conversational text. Write the final document to: {hint_path}

# EXECUTION CONSTRAINTS & WARNINGS
**Status:** WORKBOOK RESET INITIATED. Previous execution resulted in fatal errors. The workbook has been restored to its original state.
**Directive:** You must execute the original task instructions while strictly adhering to the following constraints to avoid a secondary failure.

## Fatal Pitfalls (DO NOT DO THIS)
* **[Category - e.g., Logic/Reference Error]:** [Description of the methodological error/action to avoid.]
* **[Category - e.g., Tool/Scripting Error]:** [Description of the methodological error/action to avoid.]

## Required Mitigation Strategies (MUST DO THIS)
* **[Specific Action 1]:** [Clear, actionable instruction on the correct technical approach.]
* **[Specific Action 2]:** [Clear, actionable instruction on the correct technical approach.]
"""
