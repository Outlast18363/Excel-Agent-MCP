"""Codex CLI command builder shared by all agent roles."""

import json
from pathlib import Path

EXCEL_MCP_ROOT = Path("/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp")
MODEL = "gpt-5.4-mini"
REASONING_EFFORT = "low"
APPROVAL_POLICY = "never"
SANDBOX_MODE = "workspace-write"

# Per-role excel-mcp tool allowlists. Planner only inspects; Executor mutates;
# Evaluator inspects and can trace formulas to verify the Executor's work.
_PLANNER_TOOLS = ["open_workbook", "get_sheet_state", "local_screenshot", "get_range", "close_workbook"]
ROLE_TOOLS: dict[str, list[str]] = {
    "Planner": _PLANNER_TOOLS,
    "Executor": _PLANNER_TOOLS + ["web_search", "xlwing_skills"],
    "Evaluator": _PLANNER_TOOLS + ["trace_formula"],
    # Distiller is pure text-to-text (eval report -> execution hint); no MCP
    # tools needed. Empty allowlist disables every excel-mcp tool for its subprocess.
    "Distiller": [],
}


def _excel_mcp_overrides(role: str) -> list[str]:
    """Emit the repeated `-c mcp_servers."excel-mcp".*` TOML overrides for *role*."""
    k, u = 'mcp_servers."excel-mcp".', json.dumps
    rows = (
        ("command", u(str(EXCEL_MCP_ROOT / ".venv/bin/python"))),
        ("args", u(["-m", "excel_mcp"])),
        ("cwd", u(str(EXCEL_MCP_ROOT))),
        ("startup_timeout_sec", "20"),
        ("tool_timeout_sec", "120"),
        ("enabled", "true"),
        ("enabled_tools", u(ROLE_TOOLS[role])),
    )
    return [p for n, v in rows for p in ("-c", f"{k}{n} = {v}")]


def build_codex_cmd(role: str, prompt: str, workspace: Path) -> list[str]:
    """Compose the `codex exec --json` argv for one agent invocation."""
    if role not in ROLE_TOOLS:
        raise ValueError(f"unknown role {role!r}; expected one of {list(ROLE_TOOLS)}")
    return [
        "codex",
        "exec",
        "--json",
        "--skip-git-repo-check",
        "-C", str(workspace),
        "-m", MODEL,
        *_excel_mcp_overrides(role),
        "-c", f'model_reasoning_effort="{REASONING_EFFORT}"',
        "-c", f'approval_policy="{APPROVAL_POLICY}"',
        "-s", SANDBOX_MODE,
        prompt,
    ]
