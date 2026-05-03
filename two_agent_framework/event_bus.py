"""Single-writer JSONL event bus for the two-agent workflow.

Agents run sequentially in the orchestrator, so one append-mode file handle
with flush-per-write is enough: no locking, no buffering surprises.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

_USAGE_KEYS = ("input_tokens", "cached_input_tokens", "output_tokens")


def snapshot_role_usage(role: str, usage_by_agent: dict[str, dict[str, int]]) -> dict[str, int]:
    """Copy cumulative usage counters for one agent role."""
    base = usage_by_agent.get(role) or {}
    return {k: int(base.get(k, 0) or 0) for k in _USAGE_KEYS}


def usage_delta(after: dict[str, int], before: dict[str, int]) -> dict[str, int]:
    """Per-session usage = cumulative after minus cumulative before."""
    return {k: after[k] - before[k] for k in _USAGE_KEYS}


def print_agent_session_usage(agent: str, usage: dict[str, int]) -> None:
    """Echo one session's token usage to stderr."""
    record = {"agent": agent, "type": "turn.completed", "usage": dict(usage)}
    print(json.dumps(record, ensure_ascii=False), file=sys.stderr, flush=True)


class EventBus:
    def __init__(self, jsonl_path: Path):
        self.path = Path(jsonl_path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._fh = self.path.open("a", encoding="utf-8")
        self.usage_by_agent: dict[str, dict[str, int]] = {}

    def _add_stream_usage(self, agent: str, usage: Any) -> None:
        """Sum Codex-reported usage into usage_by_agent."""
        if not isinstance(usage, dict):
            return
        acc = self.usage_by_agent.setdefault(agent, {k: 0 for k in _USAGE_KEYS})
        for key in acc:
            acc[key] += int(usage.get(key, 0) or 0)

    def emit(self, agent: str, event: dict[str, Any]) -> None:
        """Write one agent-labeled JSON line and flush immediately."""
        record = {"agent": agent, **event}
        # Historical traces use one JSON record followed by two blank lines.
        # Existing parsers skip blank lines, so keep this compatibility detail.
        self._fh.write(json.dumps(record, ensure_ascii=False) + "\n\n\n")
        self._fh.flush()
        et = event.get("type")
        if et == "turn.completed" and not event.get("session_summary"):
            self._add_stream_usage(agent, event.get("usage"))
        elif et == "turn.failed":
            self._add_stream_usage(agent, event.get("usage"))

    def agent_session_completed(
        self,
        agent: str,
        iter_idx: int,
        usage: dict[str, int],
        *,
        echo_stderr: bool = True,
    ) -> None:
        """Append the completed agent turn's token totals to the trace and stderr."""
        usage_clean = dict(usage)
        self.emit(
            agent,
            {
                "type": "turn.completed",
                "session_summary": True,
                "iter": iter_idx,
                "usage": usage_clean,
            },
        )
        if echo_stderr:
            print_agent_session_usage(agent, usage_clean)

    def transition(self, from_agent: str | None, to_agent: str, iter_idx: int, reason: str) -> None:
        self.emit("ORCHESTRATOR", {
            "type": "transition",
            "from": from_agent,
            "to": to_agent,
            "iter": iter_idx,
            "reason": reason,
        })

    def verdict(self, verdict: str, iter_idx: int) -> None:
        self.emit("Evaluator", {"type": "verdict", "verdict": verdict, "iter": iter_idx})

    def reset(self, iter_idx: int, reset_count: int, triggered_by: str) -> None:
        self.emit("ORCHESTRATOR", {
            "type": "reset",
            "iter": iter_idx,
            "reset_count": reset_count,
            "triggered_by": triggered_by,
        })

    def snapshot(
        self,
        path: Path,
        snapshot_paths: list[Path],
        source_paths: list[Path] | None = None,
    ) -> None:
        source_paths = snapshot_paths if source_paths is None else source_paths
        self.emit("ORCHESTRATOR", {
            "type": "snapshot",
            "path": str(path),
            "paths": [str(path) for path in snapshot_paths],
            "workbooks": [str(path) for path in source_paths],
        })

    def final_output_dir_wipe(self, iter_idx: int, path: Path) -> None:
        self.emit("ORCHESTRATOR", {
            "type": "final_output_dir_wipe",
            "iter": iter_idx,
            "path": str(path),
        })

    def final_workbook_files_copy(self, iter_idx: int, final_dir: Path, files: list[Path]) -> None:
        self.emit("ORCHESTRATOR", {
            "type": "final_workbook_files_copy",
            "iter": iter_idx,
            "path": str(final_dir),
            "files": [str(path) for path in files],
        })

    def restore_from_snapshot(self, iter_idx: int, snapshot_dir: Path, workbook_dir: Path) -> None:
        self.emit("ORCHESTRATOR", {
            "type": "restore_from_snapshot",
            "iter": iter_idx,
            "snapshot_dir": str(snapshot_dir),
            "workbook_dir": str(workbook_dir),
        })

    def distill(self, iter_idx: int, eval_path: Path, hint_path: Path) -> None:
        self.emit("ORCHESTRATOR", {
            "type": "distill",
            "iter": iter_idx,
            "eval_path": str(eval_path),
            "hint_path": str(hint_path),
        })

    def iteration_usage(self, iter_idx: int, usage_by_agent: dict[str, dict[str, int]]) -> None:
        self.emit("ORCHESTRATOR", {
            "type": "iteration_usage",
            "iter": iter_idx,
            "usage_by_agent": usage_by_agent,
        })

    def close(self) -> None:
        if not self._fh.closed:
            self._fh.close()

    def __enter__(self) -> "EventBus":
        return self

    def __exit__(self, *exc) -> None:
        self.close()
