"""Unit tests for BaseAgent's JSON stream parsing."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from multi_agent_framework.agent import BaseAgent
from multi_agent_framework.event_bus import EventBus


class _FakeProc:
    def __init__(self, stdout_lines: list[bytes]):
        self.stdout = iter(stdout_lines)
        self.stdin = _FakeStdin()

    def wait(self) -> int:
        return 0


class _FakeStdin:
    def write(self, _data: bytes) -> None:
        return None

    def close(self) -> None:
        return None


class _TestAgent(BaseAgent):
    ROLE = "Executor"


class BaseAgentStreamTests(unittest.TestCase):
    def test_stream_treats_json_arrays_as_raw_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            trace_path = Path(tmpdir) / "trace.jsonl"
            with EventBus(trace_path) as bus:
                agent = _TestAgent(bus, Path(tmpdir))
                proc = _FakeProc([
                    b'{"type":"item.completed","item":{"type":"agent_message","text":"hello"}}\n',
                    b'["Benchmark Return","Timing Return"]\n',
                ])

                with patch("multi_agent_framework.agent.build_codex_cmd", return_value=["codex"]), patch(
                    "multi_agent_framework.agent.subprocess.Popen",
                    return_value=proc,
                ):
                    final_msg = agent._stream("prompt")

            records = [
                line
                for line in trace_path.read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]

        self.assertEqual(final_msg, "hello")
        self.assertEqual(len(records), 2)
        self.assertIn('"type": "item.completed"', records[0])
        self.assertIn('"type": "raw"', records[1])
        self.assertIn('"line": "[\\"Benchmark Return\\",\\"Timing Return\\"]"', records[1])


if __name__ == "__main__":
    unittest.main()
