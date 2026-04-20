"""Single-writer JSONL event bus.

Agents run sequentially in the orchestrator, so one append-mode file handle
with flush-per-write is enough — no locking, no buffering surprises.
"""

import json
from pathlib import Path
from typing import Any


class EventBus:
    def __init__(self, jsonl_path: Path):
        self.path = Path(jsonl_path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._fh = self.path.open("a", encoding="utf-8")

    def emit(self, agent: str, event: dict[str, Any]) -> None:
        """Write one agent-labeled JSON line and flush immediately."""
        record = {"agent": agent, **event}
        self._fh.write(json.dumps(record, ensure_ascii=False) + "\n")
        self._fh.flush()

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

    def close(self) -> None:
        if not self._fh.closed:
            self._fh.close()

    def __enter__(self) -> "EventBus":
        return self

    def __exit__(self, *exc) -> None:
        self.close()
