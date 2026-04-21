"""Agent classes.

Each subclass has an explicit, typed `build_prompt(...)` that renders its own
f-string template, and an explicit `run(...)` that wires those same parameters
through. This makes call-site typos fail at the call itself (TypeError on
unknown/missing kwargs) rather than deep inside `str.format`.

`BaseAgent._stream` is the shared subprocess driver: it spawns
`codex exec --json`, streams events through the EventBus, and deterministically
captures the last `agent_message` item for downstream use (in particular, the
Evaluator's verdict tag).
"""

import json
import re
import subprocess
from pathlib import Path

from .config import build_codex_cmd
from .event_bus import EventBus


class BaseAgent:
    ROLE: str = ""

    def __init__(self, bus: EventBus, workspace: Path):
        self.bus = bus
        self.workspace = Path(workspace)

    def _stream(self, prompt: str) -> str:
        """Spawn codex, stream JSON events through the bus, return last agent_message text."""
        cmd = build_codex_cmd(self.ROLE, self.workspace)
        # The cmd ends with `-`, so codex reads the prompt from stdin. Piping
        # (rather than passing the prompt as an argv string) matches the
        # raw_api_baseline pattern and avoids Windows argv size/encoding issues
        # when Finch task descriptions are long.
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            bufsize=0,
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
        # Decode the child's bytes ourselves so Windows locale defaults never
        # leak into TextIOWrapper and crash the stream loop.
        for raw_line in proc.stdout:
            line = raw_line.decode("utf-8", errors="replace").rstrip("\r\n")
            if not line:
                continue
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                # non-JSON chatter (e.g. stderr merged in) — record verbatim and skip
                self.bus.emit(self.ROLE, {"type": "raw", "line": line})
                continue
            self.bus.emit(self.ROLE, event)
            # Track the most recent completed agent_message; whatever the agent
            # said last is what callers care about (verdict tag, handover note, …).
            if event.get("type") == "item.completed":
                item = event.get("item", {}) or {}
                # Codex CLI nests the item kind under "type" (not "item_type");
                # using the wrong key here silently drops every agent_message
                # and breaks the Evaluator's verdict detection.
                if item.get("type") == "agent_message":
                    final_msg = item.get("text", final_msg)
        proc.wait()
        return final_msg


class PlannerAgent(BaseAgent):
    ROLE = "Planner"

    def build_prompt(self, task: str, workbook: Path, plan_path: Path) -> str:
        return f"""\
You are the PLANNER. You do not modify the workbook. Use only inspection tools
(open_workbook, get_sheet_state, local_screenshot, get_range, close_workbook)
to understand the current state, then produce concise, structured plan another
agent can execute.

Task: {task}

Workbook (read-only for you): {workbook}

Write your final plan as Markdown to: {plan_path}
The plan must list concrete steps, target cells/ranges, formulas, and any
verification checkpoints. Here is an example:

**Task:** Fill out the Net Income column in the 'Income Statement' sheet. **Insight:** Net Income requires 'Sales Revenue' and 'Cost of Goods Sold' (COGS). 'Sales Revenue' must first be derived by backing out the 'Sales Tax' (located in the 'Cost' sheet) from the 'Sales Account' column.

**Execution Plan:**

1. Calculate the 'Sales Revenue' column using the formula: `Sales Account / (1 + Sales Tax)`. Ensure the reference to the Sales Tax cell is absolute.
2. Calculate the 'Net Income' column using the formula: `Sales Revenue - COGS`.

**Verifiable Expectations (For Evaluator):**
- **Check 1 (Absolute Referencing):** Sample 2-3 random cells in the newly calculated 'Sales Revenue' column. Verify that the formula uses an absolute reference for the Sales Tax cell (e.g., `Cost!$B$2`) so it does not shift across rows.
- **Check 2 (Row Alignment):** Sample 2-3 random cells in the 'Net Income' column (e.g., A45, A112). Verify the formula accurately references the 'Sales Revenue' and 'COGS' cells strictly without row dislocation (such as the 3rd sales revenue is added to the 4th COGS value)
"""

    def run(self, task: str, workbook: Path, plan_path: Path) -> str:
        return self._stream(self.build_prompt(task=task, workbook=workbook, plan_path=plan_path))


class ExecutorAgent(BaseAgent):
    ROLE = "Executor"

    def build_prompt(
        self,
        task: str,
        plan_path: Path,
        workbook: Path,
        final_dir: Path,
        impl_path: Path,
        impl_path_or_none: Path | None,
        eval_path_or_none: Path | None,
        hint_path_or_none: Path | None,
    ) -> str:
        return f"""\
You are the EXECUTOR. You implement the plan by modifying the workbook using
xlwing (and Excel mcp inspection tools as needed). If a prior evaluation report is
provided, treat its findings as authoritative and fix them.

Task: {task}
Plan file: {plan_path}
Prior implementation report (may be None): {impl_path_or_none}
Prior evaluation report (may be None): {eval_path_or_none}
Prior error hint (may be None): {hint_path_or_none}
Workbook to modify: {workbook}

Final deliverable folder: {final_dir}

Do not copy or export the workbook into {final_dir} — the system
copies it for you. If the task requires a non-Excel deliverable (e.g.
.txt / .md / .pdf / .docx / .csv), write it directly into {final_dir}.

Write a concise and structured implementation report (what you changed, where, and why) to:
{impl_path}
"""

    def run(
        self,
        task: str,
        plan_path: Path,
        workbook: Path,
        final_dir: Path,
        impl_path: Path,
        impl_path_or_none: Path | None,
        eval_path_or_none: Path | None,
        hint_path_or_none: Path | None,
    ) -> str:
        # The distilled hint replaces the full eval report on reset; exposing both
        # would re-leak the stale cell-level findings the Distiller exists to strip.
        if hint_path_or_none is not None and eval_path_or_none is not None:
            self.bus.emit(self.ROLE, {
                "type": "warning",
                "reason": "hint_path supplied; suppressing eval_path_or_none",
            })
            eval_path_or_none = None
        return self._stream(self.build_prompt(
            task=task,
            plan_path=plan_path,
            workbook=workbook,
            final_dir=final_dir,
            impl_path=impl_path,
            impl_path_or_none=impl_path_or_none,
            eval_path_or_none=eval_path_or_none,
            hint_path_or_none=hint_path_or_none,
        ))


class EvaluatorAgent(BaseAgent):
    ROLE = "Evaluator"

    _VERDICT_RE = re.compile(r"<verdict>\s*(success|redo|reset)\s*</verdict>", re.IGNORECASE)
    # Number of times the Evaluator will re-run itself (fresh CLI subprocess, so
    # fresh context window) when its final message is missing a verdict tag.
    MAX_VERDICT_RETRY = 2

    def build_prompt(
        self,
        plan_path: Path,
        impl_path: Path,
        workbook: Path,
        final_dir: Path,
        eval_path: Path,
        task: Path,
    ) -> str:
        return f"""\
You are the EVALUATOR. You verify the Executor's work against the plan using mcp
inspection tools and trace_formula. You do not modify the workbook.

Task: {task}
Plan file: {plan_path}
Implementation report: {impl_path}

Final deliverable folder: {final_dir}
Workbook copy (published by orchestrator, inside {final_dir}): {workbook}

Inspect the workbook copy against the Planner's expectation checks; if
{final_dir} also contains a non-xlsx file, inspect that deliverable too
and treat it as the primary answer (workbook is supporting context).

Write a Markdown evaluation report (findings, pass/fail per plan step, your verdict in one word: redo, reset, or success) to: {eval_path}

Evaluation report example:
---
**EVALUATION SUMMARY:**
[1-2 sentences summarizing the overall execution state and the primary reason for failure, if applicable.]

**EXPECTATION CHECKS:**
* **Check 1: [Name of Check from Planner] - [PASS / FAIL]**
    * *Sampled Cells:* [e.g., Cost!B45, Cost!B112]
    * *Observation:* [e.g., Formula accurately uses absolute reference `$B$2`. / Formula failed to use absolute reference; row shifted to `B46`.]
* **Check 2: [Name of Check from Planner] - [PASS / FAIL]**
    * *Sampled Cells:* [e.g., A45, A112]
    * *Observation:* [e.g., Formula correctly subtracts COGS from Sales Revenue. / Row dislocation detected: A45 references Sales Revenue from row 44.]

**ADDITIONAL ERRORS FOUND:**
[Leave blank or write "None" if no additional errors were found. Otherwise, list value, format, or visual errors using bullet points.]
* *Error 1:* [e.g., The 'Net Income' column is missing the requested currency formatting.]
* *Error 2:* [e.g., Executor accidentally overwrote the header in A1.]

**NEXT STEPS / FIX INSTRUCTIONS (If Failed):**
[If Standard Failure: Provide surgical instructions for the next Executor to fix the issues in place based on the Implementation Report.]
[If Fatal Failure: Explain exactly why a reset is necessary and list specific pitfalls the next Executor must avoid.]
---

Then, in your FINAL assistant message (not in the report file), end with
exactly one verdict tag on its own:
  <verdict>success</verdict>  — plan fully satisfied, no further work
  <verdict>redo</verdict>     — defects the Executor can fix in place
  <verdict>reset</verdict>    — Executor maded multiple interdependent mistakes, safer to roll back to the original file and start over
"""

    def run(
        self,
        plan_path: Path,
        impl_path: Path,
        workbook: Path,
        final_dir: Path,
        eval_path: Path,
        task: Path,
    ) -> tuple[str, str]:
        prompt = self.build_prompt(
            plan_path=plan_path,
            impl_path=impl_path,
            workbook=workbook,
            final_dir=final_dir,
            eval_path=eval_path,
            task=task
        )
        final_msg = ""
        for attempt in range(self.MAX_VERDICT_RETRY + 1):
            final_msg = self._stream(prompt)
            # `search` (not `match`): the tag is expected at/near the end of the message.
            m = self._VERDICT_RE.search(final_msg)
            if m:
                # regex is case-insensitive, so normalize before the orchestrator sees it.
                return final_msg, m.group(1).lower()
        # TODO(jz): post-loop fallback was truncated in the screenshot — replace
        # this placeholder with your original behavior (default verdict, raise, etc.).
        self.bus.emit(self.ROLE, {
            "type": "warning",
            "reason": f"verdict tag missing after {self.MAX_VERDICT_RETRY + 1} attempts",
            "final_msg_tail": final_msg[-200:],
        })
        return final_msg, "redo"


class DistillerAgent(BaseAgent):
    ROLE = "Distiller"

    def build_prompt(self, eval_path: Path, hint_path: Path) -> str:
        # Inlined eval report: Distiller has no MCP tools, so reading in-process
        # removes an unnecessary file-read tool call and makes the subprocess deterministic.
        eval_report = Path(eval_path).read_text(encoding="utf-8")
        return f"""\
## ROLE
You are the **Distiller Agent**. You act as the bridge between an Evaluator (who audits tasks) and an Executor (who performs tasks). You will receive the Evaluator's "Verbose Evaluation Report" below. You must process this report and generate a standalone file to guide the *next* Executor.

## OBJECTIVE
Extract the core methodological failures from the verbose evaluation report and synthesize them into strict technical guardrails. The goal is to prevent the new Executor from repeating the same logical, technical, or syntactical mistakes, without cluttering their context with irrelevant data from a file state that no longer exists.

## DISTILLATION RULES (CRITICAL)
1. **Assume a Clean Slate (No Ghost States):** The new Executor is working on a fresh, reset file. Do NOT reference specific corrupted values, specific cell errors (e.g., "Cell A15 was #REF!"), or historical artifacts from the failed run.
2. **Abstract to Methodology:** Shift from *what* went wrong to *why* it went wrong.
    * *Wrong:* "The executor put the wrong tax rate in Row 10."
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
**Directive:** You must execute the original Planner instructions while strictly adhering to the following constraints to avoid a secondary failure.

## Fatal Pitfalls (DO NOT DO THIS)
* **[Category - e.g., Logic/Reference Error]:** [Description of the methodological error/action to avoid.]
* **[Category - e.g., Tool/Scripting Error]:** [Description of the methodological error/action to avoid.]

## Required Mitigation Strategies (MUST DO THIS)
* **[Specific Action 1]:** [Clear, actionable instruction on the correct technical approach.]
* **[Specific Action 2]:** [Clear, actionable instruction on the correct technical approach.]
"""

    def run(self, eval_path: Path, hint_path: Path) -> str:
        return self._stream(self.build_prompt(eval_path=eval_path, hint_path=hint_path))
