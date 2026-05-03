"""Agent classes for the two-agent workflow.

`BaseAgent._stream` mirrors the proven multi-agent subprocess driver: it
streams `codex exec --json` output through the EventBus, preserves non-JSON
lines as raw trace events, and returns the last agent message.
"""

from __future__ import annotations

import json
import re
import subprocess
from pathlib import Path

from .config import (
    build_codex_cmd,
    twowork_subprocess_env,
)
from .event_bus import EventBus


def _raw_trace_event(line: str) -> dict:
    """Represent a non-JSON Codex output line as a trace event."""
    return {"type": "raw", "line": line}


class BaseAgent:
    ROLE: str = ""

    def __init__(self, bus: EventBus, workspace: Path, *, with_excel_mcp: bool = True):
        self.bus = bus
        self.workspace = Path(workspace)
        self.with_excel_mcp = with_excel_mcp

    def _stream(self, prompt: str) -> str:
        """Spawn codex, stream JSON events through the bus, return last agent_message text."""
        cmd = build_codex_cmd(self.ROLE, self.workspace, with_excel_mcp=self.with_excel_mcp)
        subprocess_env = twowork_subprocess_env()
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            bufsize=0,
            env=subprocess_env,
        )
        assert proc.stdin is not None
        try:
            proc.stdin.write(prompt.encode("utf-8"))
        except BrokenPipeError:
            pass
        finally:
            proc.stdin.close()

        final_msg = ""
        assert proc.stdout is not None
        for raw_line in proc.stdout:
            line = raw_line.decode("utf-8", errors="replace").rstrip("\r\n")
            if not line:
                continue
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                self.bus.emit(self.ROLE, _raw_trace_event(line))
                continue
            if not isinstance(event, dict):
                self.bus.emit(self.ROLE, _raw_trace_event(line))
                continue
            self.bus.emit(self.ROLE, event)
            if event.get("type") == "item.completed":
                item = event.get("item", {}) or {}
                if item.get("type") == "agent_message":
                    final_msg = item.get("text", final_msg)
        return_code = proc.wait()
        if return_code:
            self.bus.emit(self.ROLE, {
                "type": "process.exited",
                "return_code": return_code,
            })
        return final_msg


class WorkerAgent(BaseAgent):
    ROLE = "Worker"

    def build_prompt(
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
        if hint_path_or_none is not None and eval_path_or_none is not None:
            eval_path_or_none = None
        return f"""\
You are the WORKER. You have to implement the spreadsheet task, and you cannot end before you have completed the task.

## Required workflow
1. Before substantial edits, write an `# Execution Todo` section to {impl_path}.
2. Todo items must be actionable high-level implementation steps, not an exhaustive data inventory.
3. Mark each completed todo item with `[CHECKED]` as you complete it.
4. After implementation, complete the `# Implementation Report` section in {impl_path}.
5. Do not copy workbook/source files into {final_dir}; the orchestrator publishes staged files there after you finish.
6. If the task requires a non-xlsx deliverable, write that deliverable directly into {final_dir}.
7. ALWAYS close any workbook you open when you are done.
8. You HAVE to expliclitly calcualte the number of cells in the range you are able to query (such as calculating that A1:B100 has 200 cells). If the cell count is greater than 100, you HAVE to use the `local_screenshot` tool, or you have to let a subagent call the get range tool and return to you the summary (your prompt to the subagent should specify the specific things you are looking for).

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
- Use `local_screenshot` to quickly understand workbook layout or for range check with > 100 cells; store screenshots under {run_dir}/screenshots/.
- Use openpyxl or normal workbook edit and small range inspection and use xlwings for workbook graph complex table creation.

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
        if hint_path_or_none is not None and eval_path_or_none is not None:
            self.bus.emit(self.ROLE, {
                "type": "warning",
                "reason": "hint_path supplied; suppressing eval_path_or_none",
            })
            eval_path_or_none = None
        return self._stream(self.build_prompt(
            task=task,
            workbook_dir=workbook_dir,
            staging_note=staging_note,
            final_dir=final_dir,
            impl_path=impl_path,
            eval_path_or_none=eval_path_or_none,
            hint_path_or_none=hint_path_or_none,
            run_dir=run_dir,
        ))


class EvaluatorAgent(BaseAgent):
    ROLE = "Evaluator"

    _VERDICT_RE = re.compile(r"<verdict>\s*(success|redo|reset)\s*</verdict>", re.IGNORECASE)
    MAX_VERDICT_RETRY = 2

    def build_prompt(
        self,
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
4. When using `local_screenshot`, store screenshots under {run_dir}/screenshots/.
5. ALWAYS close any workbook you open when you are done.
6. You may use the reference file to evaluate the output of the Worker.

In your FINAL assistant message, end with exactly one verdict tag on its own:
<verdict>success</verdict>
<verdict>redo</verdict>
<verdict>reset</verdict>
"""

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
        prompt = self.build_prompt(
            task=task,
            snapshot_dir=snapshot_dir,
            final_dir=final_dir,
            impl_path=impl_path,
            eval_path=eval_path,
            staging_note=staging_note,
            run_dir=run_dir,
        )
        final_msg = ""
        for _ in range(self.MAX_VERDICT_RETRY + 1):
            final_msg = self._stream(prompt)
            match = self._VERDICT_RE.search(final_msg)
            if match:
                return final_msg, match.group(1).lower()
        self.bus.emit(self.ROLE, {
            "type": "warning",
            "reason": f"verdict tag missing after {self.MAX_VERDICT_RETRY + 1} attempts",
            "final_msg_tail": final_msg[-200:],
        })
        return final_msg, "redo"


class DistillerAgent(BaseAgent):
    ROLE = "Distiller"

    def build_prompt(self, eval_path: Path, hint_path: Path) -> str:
        eval_report = Path(eval_path).read_text(encoding="utf-8")
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

    def run(self, eval_path: Path, hint_path: Path) -> str:
        return self._stream(self.build_prompt(eval_path=eval_path, hint_path=hint_path))
