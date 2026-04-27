#!/usr/bin/env python3
"""Prompt-context regressions for orchestrator staging notes."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from multi_agent_framework.agent import EvaluatorAgent, ExecutorAgent, PlannerAgent
from multi_agent_framework.orchestrator import Orchestrator


class PromptContextTests(unittest.TestCase):
    def test_planner_prompt_includes_synthesized_staging_note(self) -> None:
        note = Orchestrator._build_staging_note([Path("workbook/72_ref_0.xlsx")], True)
        agent = PlannerAgent(bus=None, workspace=Path("."))  # type: ignore[arg-type]

        prompt = agent.build_prompt(
            task="Populate the workbook from the PDF.",
            workbook_dir=Path("workbook"),
            staging_note=note,
            plan_path=Path("handover/plan.md"),
            run_dir=Path("run"),
        )

        self.assertIn(note, prompt)
        self.assertIn("blank workbook because the task did not provide one", prompt)
        self.assertIn("Workbook folder (contains all task files, read-only for you)", prompt)

    def test_planner_prompt_includes_file_type_summary_for_normal_staging(self) -> None:
        note = Orchestrator._build_staging_note(
            [Path("workbook/3_src_0.xlsx"), Path("workbook/3_src_1.json")],
            False,
        )
        agent = PlannerAgent(bus=None, workspace=Path("."))  # type: ignore[arg-type]

        prompt = agent.build_prompt(
            task="Update the existing workbook.",
            workbook_dir=Path("workbook"),
            staging_note=note,
            plan_path=Path("handover/plan.md"),
            run_dir=Path("run"),
        )

        self.assertIn(note, prompt)
        self.assertIn(".json, .xlsx", prompt)
        self.assertNotIn("Primary workbook", prompt)

    def test_executor_and_evaluator_prompts_reference_workbook_folder(self) -> None:
        note = Orchestrator._build_staging_note(
            [Path("workbook/3_src_0.xlsx"), Path("workbook/3_src_1.csv")],
            False,
        )
        executor = ExecutorAgent(bus=None, workspace=Path("."))  # type: ignore[arg-type]
        evaluator = EvaluatorAgent(bus=None, workspace=Path("."))  # type: ignore[arg-type]

        executor_prompt = executor.build_prompt(
            task="Refresh the workbook and export a csv.",
            plan_path=Path("handover/plan.md"),
            workbook_dir=Path("workbook"),
            staging_note=note,
            final_dir=Path("final_result"),
            impl_path=Path("handover/impl_report.md"),
            impl_path_or_none=None,
            eval_path_or_none=None,
            hint_path_or_none=None,
            run_dir=Path("run"),
        )
        evaluator_prompt = evaluator.build_prompt(
            plan_path=Path("handover/plan.md"),
            impl_path=Path("handover/impl_report.md"),
            workbook_dir=Path("workbook"),
            staging_note=note,
            final_dir=Path("final_result"),
            eval_path=Path("handover/eval_report.md"),
            task="Refresh the workbook and export a csv.",
            run_dir=Path("run"),
        )

        self.assertIn("Workbook folder to modify (pick .xlsx inside it if present)", executor_prompt)
        self.assertNotIn("Workbook to modify:", executor_prompt)
        self.assertIn("Workbook folder (published by orchestrator, inside final_result): final_result", evaluator_prompt)
        self.assertNotIn("Workbook copy (published by orchestrator", evaluator_prompt)


if __name__ == "__main__":
    unittest.main()
