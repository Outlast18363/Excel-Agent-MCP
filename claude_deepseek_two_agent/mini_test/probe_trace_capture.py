from __future__ import annotations

import json
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from claude_invocation import run_claude_role, validate_child_env  # noqa: E402
from trace import EventBus  # noqa: E402


def main() -> int:
    env_error = validate_child_env()
    if env_error:
        print(env_error, file=sys.stderr)
        return 2

    run_dir = Path(__file__).resolve().parent / "probe_run"
    workbook_dir = run_dir / "workbook"
    trace_path = run_dir / "trace.jsonl"
    workbook_dir.mkdir(parents=True, exist_ok=True)
    trace_path.unlink(missing_ok=True)
    fixture_path = run_dir / "probe_fixture.txt"
    fixture_path.write_text("TRACE_TOOL_READ_FIXTURE", encoding="utf-8")

    prompt = (
        f"Use the Read tool to read this file: {fixture_path}\n"
        "Then reply with exactly this visible text:\n"
        "TRACE_VISIBLE_MESSAGE_OK"
    )
    with EventBus(trace_path) as bus:
        result = run_claude_role(
            "Distiller",
            prompt,
            bus=bus,
            run_dir=run_dir,
            workbook_dir=workbook_dir,
            with_excel_mcp=False,
        )

    records = []
    if trace_path.exists():
        for raw_line in trace_path.read_text(encoding="utf-8").splitlines():
            if raw_line.strip():
                records.append(json.loads(raw_line))

    visible_hits = [
        record for record in records
        if "TRACE_VISIBLE_MESSAGE_OK" in json.dumps(record, ensure_ascii=False)
    ]
    result_records = [record for record in records if record.get("type") == "result"]
    text_deltas = [
        record for record in records
        if record.get("type") == "stream_event"
        and ((record.get("event") or {}).get("delta") or {}).get("type") == "text_delta"
    ]
    assistant_blocks = [record for record in records if record.get("type") == "assistant_message_block"]
    tool_call_blocks = [
        record for record in assistant_blocks
        if record.get("content_type") == "tool_use"
    ]
    tool_results = [record for record in records if record.get("type") == "tool_result"]

    print(json.dumps({
        "return_code": result.return_code,
        "final_text": result.final_text,
        "usage": result.usage,
        "trace_path": str(trace_path),
        "record_count": len(records),
        "visible_hit_count": len(visible_hits),
        "result_record_count": len(result_records),
        "assistant_message_block_count": len(assistant_blocks),
        "tool_call_block_count": len(tool_call_blocks),
        "tool_result_count": len(tool_results),
        "text_delta_count": len(text_deltas),
        "first_record_types": [record.get("type") for record in records[:8]],
        "assistant_blocks": assistant_blocks,
        "tool_results": tool_results,
        "result_records": result_records,
    }, indent=2, ensure_ascii=False))
    return result.return_code


if __name__ == "__main__":
    raise SystemExit(main())
