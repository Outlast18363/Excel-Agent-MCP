from __future__ import annotations

import json
import os
import copy
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from trace import EventBus, UsageAccumulator, normalize_usage


ROOT = Path(__file__).resolve().parent
ENV_FILE = ROOT / ".env"
EXCEL_MCP_ROOT = ROOT.parent
EXCEL_MCP_PYTHON = EXCEL_MCP_ROOT / ".venv" / "Scripts" / "python.exe"

ROLE_TOOL_MASKS: dict[str, list[str]] = {
    "Worker": [
        "Read",
        "Write",
        "Edit",
        "PowerShell",
        "WebSearch",
        "WebFetch",
        "mcp__excel__open_workbook",
        "mcp__excel__local_screenshot",
        "mcp__excel__close_workbook",
        "mcp__excel__get_range",
    ],
    "Evaluator": [
        "Read",
        "Write",
        "PowerShell",
        "mcp__excel__open_workbook",
        "mcp__excel__local_screenshot",
        "mcp__excel__close_workbook",
        "mcp__excel__get_range",
    ],
    "Distiller": ["Read", "Write", "Edit"],
}


@dataclass
class ClaudeSessionResult:
    final_text: str
    usage: dict[str, int]
    return_code: int


def _ingest_thinking_start(block: dict[str, Any], content_block: dict[str, Any]) -> None:
    """Pull initial reasoning text from content_block_start (plaintext vs opaque)."""
    ct = block.get("content_type")
    if isinstance(content_block.get("thinking"), str):
        block["thinking_plain"] += content_block["thinking"]
    if isinstance(content_block.get("data"), str):
        if ct == "redacted_thinking":
            block["thinking_redacted_char_count"] += len(content_block["data"])
        else:
            block["thinking_plain"] += content_block["data"]
    if isinstance(content_block.get("signature"), str):
        block["signature_redacted_char_count"] += len(content_block["signature"])


def _ingest_thinking_delta(block: dict[str, Any], delta: dict[str, Any], delta_type: str | None) -> None:
    """Merge streaming reasoning: explicit thinking text vs opaque redacted payloads."""
    ct = block.get("content_type")
    dt = delta_type if delta_type else delta.get("type")

    if dt == "signature_delta":
        if isinstance(delta.get("signature"), str):
            block["signature_redacted_char_count"] += len(delta["signature"])
        return

    if dt == "redacted_thinking_delta":
        if isinstance(delta.get("data"), str):
            block["thinking_redacted_char_count"] += len(delta["data"])
        return

    if dt == "thinking_delta":
        if isinstance(delta.get("thinking"), str):
            block["thinking_plain"] += delta["thinking"]
        return

    # Untyped or provider-specific chunks inside a thinking block
    if isinstance(delta.get("thinking"), str):
        block["thinking_plain"] += delta["thinking"]
    if isinstance(delta.get("data"), str):
        if ct == "redacted_thinking":
            block["thinking_redacted_char_count"] += len(delta["data"])
        else:
            block["thinking_plain"] += delta["data"]
    if isinstance(delta.get("signature"), str):
        block["signature_redacted_char_count"] += len(delta["signature"])


class MessageBlockTracer:
    """Compact Claude stream blocks into readable message/tool trace records."""

    def __init__(self) -> None:
        self._blocks: dict[int, dict[str, Any]] = {}
        self._tool_names_by_id: dict[str, str] = {}

    def records_for(self, event: dict[str, Any]) -> list[dict[str, Any]]:
        if event.get("type") == "assistant":
            return []
        if event.get("type") == "user":
            return _compact_user_records(event, self._tool_names_by_id)
        if event.get("type") == "system" and event.get("subtype") == "status":
            return []
        if event.get("type") != "stream_event":
            return [_compact_trace_event(event)]

        stream = event.get("event")
        if not isinstance(stream, dict):
            return []
        stream_type = stream.get("type")
        if stream_type == "content_block_start":
            index = int(stream.get("index", -1))
            content_block = stream.get("content_block") or {}
            self._blocks[index] = {
                "index": index,
                "content_type": _content_block_type(content_block),
                "content_block": _compact_value(content_block),
                "text": "",
                "partial_json": "",
                "thinking_plain": "",
                "thinking_redacted_char_count": 0,
                "signature_redacted_char_count": 0,
            }
            if isinstance(content_block, dict):
                _ingest_thinking_start(self._blocks[index], content_block)
            return []
        if stream_type == "content_block_delta":
            index = int(stream.get("index", -1))
            block = self._blocks.setdefault(index, {
                "index": index,
                "content_type": "unknown",
                "content_block": {},
                "text": "",
                "partial_json": "",
                "thinking_plain": "",
                "thinking_redacted_char_count": 0,
                "signature_redacted_char_count": 0,
            })
            delta = stream.get("delta") or {}
            if isinstance(delta, dict):
                delta_type = delta.get("type")
                if delta_type == "text_delta":
                    block["text"] += str(delta.get("text") or "")
                elif delta_type == "input_json_delta":
                    block["partial_json"] += str(delta.get("partial_json") or "")
                elif delta_type in {"thinking_delta", "signature_delta", "redacted_thinking_delta"}:
                    _ingest_thinking_delta(block, delta, delta_type)
                elif block.get("content_type") in {"thinking", "redacted_thinking"}:
                    _ingest_thinking_delta(block, delta, delta_type)
            return []
        if stream_type == "content_block_stop":
            index = int(stream.get("index", -1))
            record = self._finish_block(index)
            return [record] if record is not None else []
        return []

    def flush(self) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []
        for index in sorted(self._blocks):
            record = self._finish_block(index)
            if record is not None:
                records.append(record)
        return records

    def _finish_block(self, index: int) -> dict[str, Any] | None:
        block = self._blocks.pop(index, None)
        if block is None:
            return None
        content_type = str(block.get("content_type") or "unknown")
        record: dict[str, Any] = {
            "type": "assistant_message_block",
            "index": int(block.get("index", index)),
            "content_type": content_type,
        }
        if content_type == "text":
            record["text"] = _truncate_text(str(block.get("text") or ""))
        elif content_type == "tool_use":
            content_block = block.get("content_block") if isinstance(block.get("content_block"), dict) else {}
            tool_use_id = str(content_block.get("id") or "")
            tool_name = str(content_block.get("name") or "")
            if tool_use_id and tool_name:
                self._tool_names_by_id[tool_use_id] = tool_name
            record["tool_name"] = tool_name or None
            tool_input = content_block.get("input")
            partial_json = str(block.get("partial_json") or "")
            if partial_json:
                tool_input = _parse_json_or_text(partial_json)
            record["tool_input"] = _compact_value(tool_input)
        elif content_type in {"thinking", "redacted_thinking"}:
            plain = str(block.get("thinking_plain") or "")
            record["thinking"] = _truncate_text(plain) if plain.strip() else ""
        else:
            record["content_block"] = _compact_value(block.get("content_block"))
            if block.get("partial_json"):
                record["partial_json"] = _truncate_text(str(block["partial_json"]))
        return record


def load_dotenv(path: Path = ENV_FILE) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("\"'")
        if key:
            values[key] = value
    return values


def claude_command() -> list[str] | None:
    if sys.platform == "win32":
        cmd_shim = shutil.which("claude.cmd")
        if cmd_shim:
            shim_path = Path(cmd_shim)
            exe_path = shim_path.parent / "node_modules" / "@anthropic-ai" / "claude-code" / "bin" / "claude.exe"
            if exe_path.is_file():
                return [str(exe_path)]
            return ["cmd", "/c", cmd_shim]
        for name in ("claude.bat", "claude"):
            found = shutil.which(name)
            if found:
                suffix = Path(found).suffix.lower()
                if suffix in {".cmd", ".bat"}:
                    return ["cmd", "/c", found]
                return [found]
        return None
    found = shutil.which("claude")
    return [found] if found else None


def child_env() -> dict[str, str]:
    env = os.environ.copy()
    dotenv = load_dotenv()
    env.update(dotenv)
    token = env.get("ANTHROPIC_AUTH_TOKEN") or env.get("ANTHROPIC_API_KEY") or env.get("DEEPSEEK_API_KEY")
    if token:
        env["ANTHROPIC_AUTH_TOKEN"] = token
        env.setdefault("ANTHROPIC_API_KEY", token)
    return env


def validate_child_env() -> str | None:
    env = child_env()
    if not (env.get("ANTHROPIC_AUTH_TOKEN") or env.get("ANTHROPIC_API_KEY")):
        return f"Missing DEEPSEEK_API_KEY, ANTHROPIC_AUTH_TOKEN, or ANTHROPIC_API_KEY in {ENV_FILE}"
    if claude_command() is None:
        return "Claude Code is not installed or not on PATH. Install it with: cmd /c npm install -g @anthropic-ai/claude-code"
    return None


def write_role_mcp_config(run_dir: Path, role: str, with_excel_mcp: bool) -> Path:
    config_dir = run_dir / "mcp_configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    path = config_dir / f"{role.lower()}_mcp.json"
    payload: dict[str, Any] = {"mcpServers": {}}
    if with_excel_mcp and role in {"Worker", "Evaluator"}:
        payload["mcpServers"]["excel"] = {
            "command": str(EXCEL_MCP_PYTHON),
            "args": ["-m", "excel_mcp"],
            "cwd": str(EXCEL_MCP_ROOT),
        }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return path


def build_claude_cmd(
    role: str,
    prompt: str,
    run_dir: Path,
    workbook_dir: Path,
    *,
    with_excel_mcp: bool,
) -> list[str]:
    command = claude_command()
    if command is None:
        raise RuntimeError("Claude Code is not installed or not on PATH.")
    mcp_config = write_role_mcp_config(run_dir, role, with_excel_mcp)
    tool_mask = ",".join(ROLE_TOOL_MASKS[role])
    return [
        *command,
        "--bare",
        "-p",
        prompt,
        "--output-format",
        "stream-json",
        "--verbose",
        "--include-partial-messages",
        "--include-hook-events",
        "--strict-mcp-config",
        "--mcp-config",
        str(mcp_config),
        "--permission-mode",
        "bypassPermissions",
        "--allowedTools",
        tool_mask,
        "--add-dir",
        str(workbook_dir),
    ]


def _runner_popen_kwargs() -> dict[str, object]:
    if sys.platform == "win32":
        return {"creationflags": getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)}
    return {"start_new_session": True}


def terminate_process_tree(process: subprocess.Popen[str]) -> None:
    if process.poll() is not None:
        return
    if sys.platform == "win32":
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        return
    process.terminate()


def _extract_text_from_event(event: dict[str, Any], current: str) -> str:
    if event.get("type") == "result" and isinstance(event.get("result"), str):
        return event["result"]
    if event.get("type") == "assistant" and isinstance(event.get("message"), dict):
        return _extract_text_from_message(event["message"], current)
    if event.get("type") == "stream_event":
        stream = event.get("event")
        if isinstance(stream, dict):
            if stream.get("type") == "content_block_delta":
                delta = stream.get("delta") or {}
                if isinstance(delta, dict) and delta.get("type") == "text_delta":
                    return current + str(delta.get("text") or "")
            if stream.get("type") == "message_stop":
                message = stream.get("message")
                if isinstance(message, dict):
                    return _extract_text_from_message(message, current)
    return current


def _extract_text_from_message(message: dict[str, Any], fallback: str) -> str:
    blocks = message.get("content")
    if not isinstance(blocks, list):
        return fallback
    parts: list[str] = []
    for block in blocks:
        if isinstance(block, dict) and block.get("type") == "text":
            parts.append(str(block.get("text") or ""))
    return "".join(parts) or fallback


def run_claude_role(
    role: str,
    prompt: str,
    *,
    bus: EventBus,
    run_dir: Path,
    workbook_dir: Path,
    with_excel_mcp: bool,
) -> ClaudeSessionResult:
    cmd = build_claude_cmd(role, prompt, run_dir, workbook_dir, with_excel_mcp=with_excel_mcp)
    env = child_env()
    accumulator = UsageAccumulator()
    block_tracer = MessageBlockTracer()
    final_text = ""

    bus.emit(role, {
        "type": "process.started",
        "tool_mask": ROLE_TOOL_MASKS[role],
        "with_excel_mcp": bool(with_excel_mcp),
    })

    try:
        proc = subprocess.Popen(
            cmd,
            cwd=str(run_dir),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            **_runner_popen_kwargs(),
        )
    except OSError as exc:
        bus.emit(role, {"type": "process.failed_to_start", "error": str(exc)})
        raise

    assert proc.stdout is not None
    for raw_line in proc.stdout:
        line = raw_line.rstrip("\r\n")
        if not line:
            continue
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            bus.emit(role, {"type": "raw", "line": line})
            continue
        if not isinstance(event, dict):
            bus.emit(role, {"type": "raw", "line": line})
            continue
        for trace_record in block_tracer.records_for(event):
            bus.emit(role, trace_record)
        accumulator.observe(event)
        final_text = _extract_text_from_event(event, final_text)

    return_code = proc.wait()
    for trace_record in block_tracer.flush():
        bus.emit(role, trace_record)
    usage = normalize_usage(accumulator.total())
    bus.add_usage(role, usage)
    bus.emit(role, {"type": "process.exited", "return_code": return_code, "usage": usage})
    if return_code:
        bus.emit(role, {"type": "turn.failed", "return_code": return_code, "usage": usage})
    return ClaudeSessionResult(final_text=final_text, usage=usage, return_code=return_code)


def _sanitize_trace_event(event: dict[str, Any]) -> dict[str, Any]:
    return _sanitize_obj(event)


def _compact_trace_event(event: dict[str, Any]) -> dict[str, Any]:
    event_type = event.get("type")
    if event_type == "system":
        record = {"type": "system"}
        for key in ("subtype", "status", "model", "permissionMode"):
            if key in event:
                record[key] = event[key]
        if isinstance(event.get("tools"), list):
            record["tools"] = event["tools"]
        if isinstance(event.get("mcp_servers"), list):
            record["mcp_servers"] = event["mcp_servers"]
        return record
    if event_type == "result":
        record = {"type": "result"}
        if isinstance(event.get("result"), str):
            record["result"] = event["result"]
        if isinstance(event.get("usage"), dict):
            record["usage"] = normalize_usage(event["usage"])
        for key in ("is_error", "subtype", "total_cost_usd", "duration_ms"):
            if key in event:
                record[key] = event[key]
        return record
    if event_type in {"hook_event", "api_error", "api_retry"}:
        return _sanitize_trace_event(event)
    return {"type": str(event_type or "event")}


def _compact_user_records(event: dict[str, Any], tool_names_by_id: dict[str, str]) -> list[dict[str, Any]]:
    message = event.get("message")
    if not isinstance(message, dict):
        return []
    content = message.get("content")
    if not isinstance(content, list):
        return []
    records: list[dict[str, Any]] = []
    for item in content:
        if not isinstance(item, dict) or item.get("type") != "tool_result":
            continue
        tool_use_id = str(item.get("tool_use_id") or "")
        records.append({
            "type": "tool_result",
            "tool_name": tool_names_by_id.get(tool_use_id) or "unknown",
            "is_error": bool(item.get("is_error", False)),
            "content": _compact_tool_result_content(item.get("content")),
        })
    return records


def _compact_tool_result_content(content: Any) -> Any:
    if isinstance(content, str):
        return {"type": "text", "text": _truncate_text(content)}
    if isinstance(content, list):
        return [_compact_tool_result_content(item) for item in content[:8]]
    if isinstance(content, dict):
        content_type = content.get("type")
        if content_type in {"image", "image_url"} or "base64" in content or "data" in content:
            return {"type": str(content_type or "binary"), "omitted": True}
        return _compact_value(content)
    return _compact_value(content)


def _compact_value(value: Any, *, max_string: int = 4000) -> Any:
    value = _sanitize_obj(value)
    if isinstance(value, str):
        return _truncate_text(value, max_string)
    if isinstance(value, list):
        compacted = [_compact_value(item, max_string=max_string) for item in value[:20]]
        if len(value) > 20:
            compacted.append({"omitted_items": len(value) - 20})
        return compacted
    if isinstance(value, dict):
        compacted: dict[str, Any] = {}
        for index, (key, item) in enumerate(value.items()):
            if index >= 40:
                compacted["omitted_keys"] = len(value) - 40
                break
            compacted[str(key)] = _compact_value(item, max_string=max_string)
        return compacted
    return value


def _truncate_text(text: str, limit: int = 4000) -> str:
    if len(text) <= limit:
        return text
    return f"{text[:limit]}...<truncated {len(text) - limit} chars>"


def _parse_json_or_text(raw: str) -> Any:
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return raw


def _content_block_type(content_block: Any) -> str:
    if isinstance(content_block, dict):
        return str(content_block.get("type") or "unknown")
    return "unknown"


def _sanitize_obj(value: Any) -> Any:
    return copy.deepcopy(value)
