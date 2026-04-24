#!/usr/bin/env python3
"""Prompt-context regressions for orchestrator workbook notes."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from multi_agent_framework.agent import PlannerAgent
from multi_agent_framework.orchestrator import Orchestrator


class PromptContextTests(unittest.TestCase):
    def test_planner_prompt_includes_synthesized_workbook_note(self) -> None:
        note = Orchestrator._build_workbook_note(True)
        agent = PlannerAgent(bus=None, workspace=Path("."))  # type: ignore[arg-type]

        prompt = agent.build_prompt(
            task="Populate the workbook from the PDF.",
            workbook=Path("workbook/72_ref_0.xlsx"),
            workbook_dir=Path("workbook"),
            workbook_note=note,
            plan_path=Path("handover/plan.md"),
            run_dir=Path("run"),
        )

        self.assertIn(note, prompt)
        self.assertIn("blank workbook because the task did not provide one", prompt)

    def test_planner_prompt_includes_normal_workbook_note(self) -> None:
        note = Orchestrator._build_workbook_note(False)
        agent = PlannerAgent(bus=None, workspace=Path("."))  # type: ignore[arg-type]

        prompt = agent.build_prompt(
            task="Update the existing workbook.",
            workbook=Path("workbook/3_src_0.xlsx"),
            workbook_dir=Path("workbook"),
            workbook_note=note,
            plan_path=Path("handover/plan.md"),
            run_dir=Path("run"),
        )

        self.assertIn(note, prompt)
        self.assertIn("came from the staged task files", prompt)


if __name__ == "__main__":
    unittest.main()
