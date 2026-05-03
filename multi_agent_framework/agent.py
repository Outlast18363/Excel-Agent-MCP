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
            if not isinstance(event, dict):
                # Some child commands print JSON payloads like arrays. Those are
                # valid JSON, but they are not Codex event objects and should be
                # preserved as raw output instead of crashing EventBus.emit().
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

    def build_prompt(
        self,
        task: str,
        workbook_dir: Path,
        staging_note: str,
        plan_path: Path,
        run_dir: Path,
    ) -> str:
        return f"""\
You are the INFO_RETRIEVER. You do not modify the workbook. Use only inspection tools to produce a report with a curated list of potential items that may be relavent to the task. You should include a concise list of excel item names, and indicate if they are columns, rows, or other, with appropreate address indication (e.g. A1, A:A, A1:A10, etc.). If the task contains non xlsx source files, you may use suitable libraries to inspect them and include a breif description of important file structure and content. No need to read the excel file if its the orchetrator generated empty workbook. Important: you may write to the plan file incrementally as you read through the workbook. You don't need to write the whole plan in one go.

## Behavoir Requirements:
1. You may try to use local_screenshot tool to get a visual of the workbook, and use it to help you quickly understand the workbook structure and the header information.

2. Prefer local openpyxl inspection for simple read-only verification; remember to use MCP tools when you need screenshots or cell value search for more efficient search.

3. Your plan should be concise and structured.

4. When using the local_screenshot tool, always store the screenshots in the screenshots folder within the run dir: {run_dir}/screenshots/.

5. ALWAYS close the workbook when you are done with it.

6. You should conclude your plan when you feel like what you have written is enough to guide the executor to solve the task, DON'T overflow your plan file withn surplus information.

7. Your report shouldn't contain opionion, you should only include factual information.

## Important Information:
Task: {task}

Workbook folder (contains all task files, read-only for you): {workbook_dir}
{staging_note}

Write your final plan as Markdown to: {plan_path}

"""

    def run(
        self,
        task: str,
        workbook_dir: Path,
        staging_note: str,
        plan_path: Path,
        run_dir: Path,
    ) -> str:
        return self._stream(self.build_prompt(
            task=task,
            workbook_dir=workbook_dir,
            staging_note=staging_note,
            plan_path=plan_path,
            run_dir=run_dir,
        ))


class ExecutorAgent(BaseAgent):
    ROLE = "Executor"

    def build_prompt(
        self,
        task: str,
        plan_path: Path,
        workbook_dir: Path,
        staging_note: str,
        final_dir: Path,
        impl_path: Path,
        impl_path_or_none: Path | None,
        eval_path_or_none: Path | None,
        hint_path_or_none: Path | None,
        run_dir: Path,
    ) -> str:
        return f"""\
You are the EXECUTOR. Your job is to execute the task using the plan file as the authoritative source of workbook structure.

The plan file is the only approved source of workbook discovery. You must not inspect the workbook to rediscover structure, search for labels, scan for blank areas, hunt for templates, verify the planner's claims, or choose between alternative source ranges on your own.

Before editing, write a concise implementation checklist to {impl_path}. Every checklist item must reference exact workbook addresses from the plan file and must be a specific executable step.

If a prior evaluation report is provided, treat its findings as authoritative and fix them directly.

Rules:
1. Do not copy or export the workbook into {final_dir}; the system copies it for you. If the task requires a non-Excel deliverable, write it directly into {final_dir}.
2. Do not perform exploratory workbook inspection of any kind. Forbidden actions include `open_workbook`, `get_range`, `search_cell`, `get_sheet_state`, `local_screenshot`, openpyxl scans, xlwings scans, row/column dumps, pattern searches, and any other workbook read whose purpose is discovery rather than execution.
3. You may read workbook cells only when all of the following are true:
   - the exact sheet and exact range are already explicitly named in the plan file or evaluation report
   - the read is strictly necessary to compute or write the requested result
   - the read is limited to that exact range
4. You may verify workbook contents only after editing, and only for the exact cells you wrote or the exact cells named in the evaluation report. No neighboring or broad verification reads are allowed.
5. If the plan file does not provide an exact range or exact output area required to complete the task, do not inspect the workbook to fill the gap. Instead, write a brief note in {impl_path} that the plan is insufficient for safe execution and stop.
6. Do not resolve workbook ambiguities yourself. If the plan file lists alternatives, follow the evaluation report if one exists. Otherwise, document the blocking ambiguity in {impl_path} and stop.
7. Use xlwings for workbook edits and graph creation when editing is required.
8. BEFORE any standalone xlwings/openpyxl edit, close any MCP workbook session. Edit with exclusive access, save, close Excel, then reopen only if an exact-cell verification is required by rule 4.
9. If the task is about graph construction, you may use `local_screenshot` only to verify the exact graph or exact output area you created. Store screenshots in `{run_dir}/screenshots/`.
10. After implementation, append a brief `Implementation Report` section to {impl_path}.
11. Mark completed checklist items with `[CHECKED]`.
12. ALWAYS close the workbook when you are done.

Do not break these rules even if reinspection seems helpful. Your role is execution only, not workbook discovery.

## Important Information:
Task: {task}
Plan file: {plan_path}
Prior implementation report (may be None): {impl_path_or_none}
Prior evaluation report (may be None): {eval_path_or_none}
Prior error hint (may be None): {hint_path_or_none}
Workbook folder to modify (pick .xlsx inside it if present): {workbook_dir}
{staging_note}
Final deliverable folder: {final_dir}
"""

    def run(
        self,
        task: str,
        plan_path: Path,
        workbook_dir: Path,
        staging_note: str,
        final_dir: Path,
        impl_path: Path,
        impl_path_or_none: Path | None,
        eval_path_or_none: Path | None,
        hint_path_or_none: Path | None,
        run_dir: Path,
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
            workbook_dir=workbook_dir,
            staging_note=staging_note,
            final_dir=final_dir,
            impl_path=impl_path,
            impl_path_or_none=impl_path_or_none,
            eval_path_or_none=eval_path_or_none,
            hint_path_or_none=hint_path_or_none,
            run_dir=run_dir,
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
        workbook_dir: Path,
        staging_note: str,
        final_dir: Path,
        eval_path: Path,
        task: str,
        run_dir: Path,
    ) -> str:
        return f"""\

## Role Description:        
You are the EVALUATOR. You verify the Executor Agent's work through its final deliverable {final_dir} and the implementation report {impl_path} based on the task description {task}. The before-edit workbook folder is under the 'snapshot' folder in {workbook_dir}. It is READ ONLY. 

Staging note: 
{staging_note}

You need to read through the provided files first and come up with a checklist of checks to verify whether the executor correctly fulfilled the task to write to your evaluation report at {eval_path}. Then, you may check against the checklist one by one (mark the checklist items in the report as PASSED or FAILED with concise yet informative explainations as you go through the checklist items). Finally, provide a one word final verdict in a "Verdict" sectionat the end of your evaluation report to indicate whether the executor successfully fulfilled the task. The verdict should be one of the following: "success", "redo", or "reset". Here is the standard for each verdict:

success: task is fully satisfied, no further work.
redo: defects in implementation that the Executor can fix in place in the current workbook file.
reset: Executor maded multiple interdependent mistakes, safer to roll back to the original workbook file and start over.

## Behavoir Requirements:

1. If there are non xlsx files in the final deliverable folder {final_dir}, inspect that deliverable and treat it as the answer from the executor (workbook is supporting context).

2. When using the local_screenshot tool, always store the screenshots in the screenshots folder within the run dir: {run_dir}/screenshots/.

3. ALWAYS close the workbook when you are done with it.

## Evaluation report example:
---
**EVALUATION SUMMARY:**
[1-2 sentences summarizing the overall execution state and the primary reason for failure, if applicable.]

**EXPECTATION CHECKS:**
* **Check 1: [Name of Check from Planner] - [PASS / FAIL]**
    * *Cells involved:*
    * *Observation:* 
* **Check 2: [Name of Check from Planner] - [PASS / FAIL]**
    * *Cells involved:*
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
exactly one verdict tag on its own (same as the verdict in your evaluation report):
  <verdict>success</verdict>
  <verdict>redo</verdict> 
  <verdict>reset</verdict>
"""

    def run(
        self,
        plan_path: Path,
        impl_path: Path,
        workbook_dir: Path,
        staging_note: str,
        final_dir: Path,
        eval_path: Path,
        task: str,
        run_dir: Path,
    ) -> tuple[str, str]:
        prompt = self.build_prompt(
            plan_path=plan_path,
            impl_path=impl_path,
            workbook_dir=workbook_dir,
            staging_note=staging_note,
            final_dir=final_dir,
            eval_path=eval_path,
            task=task,
            run_dir=run_dir,
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
