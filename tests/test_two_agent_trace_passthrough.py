"""Tests for two-agent trace payload passthrough."""

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

from two_agent_framework.agent import BaseAgent, _raw_trace_event
from two_agent_framework.event_bus import EventBus


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
    ROLE = "Worker"


class TwoAgentTracePassthroughTests(unittest.TestCase):
    def test_emit_records_turn_usage_by_agent(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            trace_path = Path(tmpdir) / "trace.jsonl"
            with EventBus(trace_path) as bus:
                bus.emit("Worker", {
                    "type": "turn.completed",
                    "usage": {
                        "input_tokens": 3,
                        "cached_input_tokens": 1,
                        "output_tokens": 5,
                    },
                })
                bus.agent_session_completed(
                    "Worker",
                    0,
                    {
                        "input_tokens": 3,
                        "cached_input_tokens": 1,
                        "output_tokens": 5,
                    },
                    echo_stderr=False,
                )
                usage_by_agent = dict(bus.usage_by_agent)

        self.assertEqual(usage_by_agent["Worker"]["input_tokens"], 3)
        self.assertEqual(usage_by_agent["Worker"]["cached_input_tokens"], 1)
        self.assertEqual(usage_by_agent["Worker"]["output_tokens"], 5)

    def test_emit_preserves_nested_error_message(self) -> None:
        message = "FRONT_ONLY:" + ("x" * 5000) + "final traceback line"

        with tempfile.TemporaryDirectory() as tmpdir:
            trace_path = Path(tmpdir) / "trace.jsonl"
            with EventBus(trace_path) as bus:
                bus.emit("Worker", {"type": "turn.failed", "error": {"message": message}})
            record = json.loads(next(line for line in trace_path.read_text().splitlines() if line))

        self.assertEqual(record["error"]["message"], message)
        self.assertNotIn("message_trace_truncated", record["error"])

    def test_stream_returns_full_agent_message_text(self) -> None:
        huge_text = "FRONT_ONLY:" + ("x" * 5000) + "final verdict\n<verdict>success</verdict>"
        raw_event = json.dumps({
            "type": "item.completed",
            "item": {"id": "item_1", "type": "agent_message", "text": huge_text},
        }).encode("utf-8") + b"\n"

        with tempfile.TemporaryDirectory() as tmpdir:
            trace_path = Path(tmpdir) / "trace.jsonl"
            with EventBus(trace_path) as bus:
                agent = _TestAgent(bus, Path(tmpdir), with_excel_mcp=False)
                proc = _FakeProc([raw_event])
                with patch("two_agent_framework.agent.build_codex_cmd", return_value=["codex"]), patch(
                    "two_agent_framework.agent.twowork_subprocess_env",
                    return_value={},
                ), patch(
                    "two_agent_framework.agent.subprocess.Popen",
                    return_value=proc,
                ):
                    final_msg = agent._stream("prompt")
            record = json.loads(next(line for line in trace_path.read_text().splitlines() if line))

        self.assertEqual(final_msg, huge_text)
        self.assertEqual(record["item"]["text"], huge_text)
        self.assertNotIn("text_trace_truncated", record["item"])

    def test_raw_trace_event_preserves_long_lines(self) -> None:
        line = "FRONT_ONLY:" + ("a" * 2500) + "TRACEBACK_END"

        event = _raw_trace_event(line)

        self.assertEqual(event["type"], "raw")
        self.assertEqual(event["line"], line)
        self.assertNotIn("suppressed_keep", event)


if __name__ == "__main__":
    unittest.main()
