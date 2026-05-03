from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any


USAGE_KEYS = (
    "input_tokens",
    "cached_input_tokens",
    "cache_creation_input_tokens",
    "cache_read_input_tokens",
    "output_tokens",
)


def empty_usage() -> dict[str, int]:
    return {key: 0 for key in USAGE_KEYS}


def normalize_usage(usage: Any) -> dict[str, int]:
    clean = empty_usage()
    if not isinstance(usage, dict):
        return clean
    for key in USAGE_KEYS:
        value = usage.get(key, 0)
        if value is None:
            continue
        try:
            clean[key] = int(value)
        except (TypeError, ValueError):
            clean[key] = 0

    # Some stream producers use the older combined cached-input key. Keep both
    # detailed cache counters when present, but also preserve old summary shape.
    if not clean["cached_input_tokens"]:
        clean["cached_input_tokens"] = clean["cache_creation_input_tokens"] + clean["cache_read_input_tokens"]
    return clean


def usage_add(left: dict[str, int], right: dict[str, int]) -> dict[str, int]:
    return {key: int(left.get(key, 0) or 0) + int(right.get(key, 0) or 0) for key in USAGE_KEYS}


def usage_delta(after: dict[str, int], before: dict[str, int]) -> dict[str, int]:
    return {key: int(after.get(key, 0) or 0) - int(before.get(key, 0) or 0) for key in USAGE_KEYS}


def snapshot_role_usage(role: str, usage_by_agent: dict[str, dict[str, int]]) -> dict[str, int]:
    return normalize_usage(usage_by_agent.get(role) or {})


class UsageAccumulator:
    """Best-effort token accounting for Claude Code stream-json events."""

    def __init__(self) -> None:
        self._message_usages: list[dict[str, int]] = []
        self._result_usage: dict[str, int] | None = None

    def observe(self, event: dict[str, Any]) -> None:
        direct_usage = event.get("usage")
        if event.get("type") == "result" and isinstance(direct_usage, dict):
            self._result_usage = normalize_usage(direct_usage)
            return

        if isinstance(direct_usage, dict) and event.get("type") in {"assistant", "message"}:
            self._message_usages.append(normalize_usage(direct_usage))
            return

        if event.get("type") != "stream_event":
            return
        stream = event.get("event")
        if not isinstance(stream, dict):
            return
        stream_type = stream.get("type")
        if stream_type == "message_start":
            message = stream.get("message") or {}
            if isinstance(message, dict) and isinstance(message.get("usage"), dict):
                self._message_usages.append(normalize_usage(message["usage"]))
        elif stream_type == "message_delta" and isinstance(stream.get("usage"), dict):
            # Some SDK streams report final output usage on message_delta.
            # Treat it as a message contribution only when no result summary
            # arrives later.
            self._message_usages.append(normalize_usage(stream["usage"]))

    def total(self) -> dict[str, int]:
        if self._result_usage is not None:
            return dict(self._result_usage)
        total = empty_usage()
        for usage in self._message_usages:
            total = usage_add(total, usage)
        return total


class EventBus:
    def __init__(self, jsonl_path: Path):
        self.path = Path(jsonl_path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._fh = self.path.open("a", encoding="utf-8")
        self.usage_by_agent: dict[str, dict[str, int]] = {}

    def emit(self, agent: str, event: dict[str, Any]) -> None:
        record = {"agent": agent, **event}
        self._fh.write(json.dumps(record, ensure_ascii=False) + "\n\n")
        self._fh.flush()

    def add_usage(self, agent: str, usage: dict[str, int]) -> None:
        current = normalize_usage(self.usage_by_agent.get(agent) or {})
        self.usage_by_agent[agent] = usage_add(current, normalize_usage(usage))

    def agent_session_completed(self, agent: str, iter_idx: int, usage: dict[str, int]) -> None:
        record = {
            "type": "turn.completed",
            "session_summary": True,
            "iter": iter_idx,
            "usage": normalize_usage(usage),
        }
        self.emit(agent, record)
        print(json.dumps({"agent": agent, **record}, ensure_ascii=False), file=sys.stderr, flush=True)

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

    def snapshot(self, path: Path, snapshot_paths: list[Path], source_paths: list[Path]) -> None:
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

    def __exit__(self, *exc: object) -> None:
        self.close()
