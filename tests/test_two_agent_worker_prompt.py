"""Worker prompt regressions for required implementation-report side effects."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from two_agent_framework.agent import WorkerAgent


class WorkerPromptTests(unittest.TestCase):
    def test_worker_prompt_requires_tool_call_before_progress_message(self) -> None:
        agent = WorkerAgent(bus=None, workspace=Path("."), with_excel_mcp=False)  # type: ignore[arg-type]

        prompt = agent.build_prompt(
            task="{}",
            workbook_dir=Path("workbook"),
            staging_note="Staging note.",
            final_dir=Path("final_result"),
            impl_path=Path("handover/impl_report.md"),
            eval_path_or_none=None,
            hint_path_or_none=None,
            run_dir=Path("run"),
        )

        self.assertIn(
            "your first action after reasoning MUST be a shell command or file-edit tool call",
            prompt,
        )
        self.assertIn("not an assistant progress message", prompt)
        self.assertIn(
            "Do not send any assistant message until after handover\\impl_report.md exists".replace("\\", "/"),
            prompt.replace("\\", "/"),
        )


if __name__ == "__main__":
    unittest.main()
