#!/usr/bin/env python3
"""Cleanup regressions for reused orchestrator run directories."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from multi_agent_framework.orchestrator import Orchestrator


class OrchestratorCleanupTests(unittest.TestCase):
    def test_reused_run_dir_clears_stale_screenshots_but_keeps_staged_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            run_dir = root / "0"
            workbook_dir = run_dir / "workbook"
            screenshots_dir = run_dir / "screenshots"
            handover_dir = run_dir / "handover"
            final_dir = run_dir / "final_result"

            workbook_dir.mkdir(parents=True)
            screenshots_dir.mkdir(parents=True)
            handover_dir.mkdir(parents=True)
            final_dir.mkdir(parents=True)

            staged_workbook = workbook_dir / "0_src_0.xlsx"
            stale_workbook = workbook_dir / "old_notes.txt"
            stale_screenshot = screenshots_dir / "income_statement_top.png"
            stale_plan = handover_dir / "plan.md"
            stale_final = final_dir / "0_old_result.xlsx"

            staged_workbook.write_bytes(b"current workbook")
            stale_workbook.write_text("stale workbook artifact", encoding="utf-8")
            stale_screenshot.write_bytes(b"stale screenshot")
            stale_plan.write_text("old plan", encoding="utf-8")
            stale_final.write_bytes(b"old final")

            orchestrator = Orchestrator(
                task="Refresh workbook values",
                workbooks=[staged_workbook],
                empty_workbook_created=False,
                run_dir=run_dir,
                task_id="0",
            )

            self.assertEqual(orchestrator.screenshots_dir, screenshots_dir.resolve())
            self.assertTrue(staged_workbook.exists())
            self.assertFalse(stale_workbook.exists())
            self.assertTrue(screenshots_dir.exists())
            self.assertEqual(list(screenshots_dir.iterdir()), [])
            self.assertEqual(list(handover_dir.iterdir()), [])
            self.assertEqual(list(final_dir.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
