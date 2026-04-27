#!/usr/bin/env python3
"""Regression tests for workbook-folder anchored orchestration."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from multi_agent_framework.orchestrator import Orchestrator


def _read_trace_records(path: Path) -> list[dict]:
    return [
        json.loads(line)
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]


class OrchestratorWorkbookDirTests(unittest.TestCase):
    def test_orchestrator_accepts_non_excel_workbook_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            run_dir = root / "17"
            workbook_dir = run_dir / "workbook"
            workbook_dir.mkdir(parents=True)

            (workbook_dir / "17_src_0.csv").write_text("month,value\nJan,10\n", encoding="utf-8")
            (workbook_dir / "17_src_1.json").write_text('{"status":"ok"}\n', encoding="utf-8")

            prompts: dict[str, str] = {}

            def fake_stream(agent, prompt: str) -> str:
                prompts[agent.ROLE] = prompt
                if agent.ROLE == "Evaluator":
                    return "<verdict>success</verdict>"
                return "ok"

            with patch("multi_agent_framework.agent.BaseAgent._stream", new=fake_stream):
                result = Orchestrator(
                    task="Inspect staged csv and json files.",
                    workbook_dir=workbook_dir,
                    empty_workbook_created=False,
                    run_dir=run_dir,
                    task_id="17",
                ).run()

            self.assertEqual(result.verdict, "success")
            self.assertEqual(
                sorted(path.name for path in (run_dir / "snapshots").iterdir()),
                ["17_src_0.csv", "17_src_1.json"],
            )
            self.assertEqual(
                sorted(path.name for path in (run_dir / "final_result").iterdir()),
                ["17_src_0.csv", "17_src_1.json"],
            )
            self.assertIn(str(workbook_dir.resolve()), prompts["Planner"])
            self.assertIn("Staging note: staged task files include .csv, .json.", prompts["Planner"])
            self.assertNotIn("Primary workbook", prompts["Planner"])

            snapshot_events = [
                record
                for record in _read_trace_records(run_dir / "trace.jsonl")
                if record.get("agent") == "ORCHESTRATOR" and record.get("type") == "snapshot"
            ]
            self.assertEqual(len(snapshot_events), 1)
            self.assertEqual(snapshot_events[0]["path"], str((run_dir / "snapshots").resolve()))
            self.assertEqual(
                sorted(Path(path).name for path in snapshot_events[0]["workbooks"]),
                ["17_src_0.csv", "17_src_1.json"],
            )

    def test_orchestrator_keeps_mixed_file_runs_compatible(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            run_dir = root / "4"
            workbook_dir = run_dir / "workbook"
            workbook_dir.mkdir(parents=True)

            staged_workbook = workbook_dir / "4_src_0.xlsx"
            staged_companion = workbook_dir / "4_src_1.json"
            staged_workbook.write_bytes(b"placeholder workbook bytes")
            staged_companion.write_text('{"sheet":"Inputs"}\n', encoding="utf-8")

            prompts: dict[str, str] = {}

            def fake_stream(agent, prompt: str) -> str:
                prompts[agent.ROLE] = prompt
                if agent.ROLE == "Evaluator":
                    return "<verdict>success</verdict>"
                return "ok"

            orchestrator = Orchestrator(
                task="Update the workbook using the companion json file.",
                workbooks=[staged_workbook, staged_companion],
                empty_workbook_created=False,
                run_dir=run_dir,
                task_id="4",
            )
            self.assertEqual([path.name for path in orchestrator.xlsx_files], ["4_src_0.xlsx"])

            with patch("multi_agent_framework.agent.BaseAgent._stream", new=fake_stream):
                result = orchestrator.run()

            self.assertEqual(result.verdict, "success")
            self.assertEqual(
                sorted(path.name for path in (run_dir / "snapshots").iterdir()),
                ["4_src_0.xlsx", "4_src_1.json"],
            )
            self.assertEqual(
                sorted(path.name for path in (run_dir / "final_result").iterdir()),
                ["4_src_0.xlsx", "4_src_1.json"],
            )
            self.assertIn(
                "Workbook folder to modify (pick .xlsx inside it if present)",
                prompts["Executor"],
            )
            self.assertIn("Staging note: staged task files include .json, .xlsx.", prompts["Executor"])
            self.assertNotIn("Workbook to modify:", prompts["Executor"])


if __name__ == "__main__":
    unittest.main()
