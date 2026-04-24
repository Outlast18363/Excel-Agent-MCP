"""Codex CLI command builder shared by all agent roles."""

import json
import os
import shutil
import sys
from pathlib import Path

CODEX_CLI = (
    os.environ.get("CODEX_CLI")
    or os.environ.get("CODEX_BINARY")
    or shutil.which("codex")
    or "codex"
)
EXCEL_MCP_ROOT = Path(os.environ.get("EXCEL_MCP_ROOT") or Path(__file__).resolve().parents[1])
MODEL = "gpt-5.4"
REASONING_EFFORT = "high"
SERVICE_TIER = "fast"
FAST_MODE = True
APPROVAL_POLICY = "never"
# Fast tier is CLI-accepted here, but Codex docs say it only takes effect for ChatGPT-authenticated sessions, not API-key logins.
# On Windows, codex's `workspace-write` sandbox has been observed to cause native
# sandbox access failures (documented in raw_api_baseline/README.md). Match the
# raw_api_baseline default there: `danger-full-access` paired with the
# `--dangerously-bypass-approvals-and-sandbox` flag below. macOS/Linux keep the
# original `workspace-write` behavior the framework was written for.
SANDBOX_MODE = "danger-full-access" if os.name == "nt" else "workspace-write"

# Per-role excel-mcp tool allowlists. Planner only inspects; Executor mutates;
# Evaluator inspects and can trace formulas to verify the Executor's work.
_PLANNER_TOOLS = ["open_workbook", "get_sheet_state", "local_screenshot", "get_range", "close_workbook", "search_cell"]
ROLE_TOOLS: dict[str, list[str]] = {
    "Planner": _PLANNER_TOOLS,
    "Executor": _PLANNER_TOOLS + ["web_search"],
    "Evaluator": _PLANNER_TOOLS + ["trace_formula"],
    # Distiller is pure text-to-text (eval report -> execution hint); no MCP
    # tools needed. Empty allowlist disables every excel-mcp tool for its subprocess.
    "Distiller": [],
}


def _venv_python(root: Path) -> Path:
    """Return the virtualenv Python path for the current platform."""
    if sys.platform.startswith("win"):
        return root / ".venv" / "Scripts" / "python.exe"
    return root / ".venv" / "bin" / "python"


def _excel_mcp_overrides(role: str) -> list[str]:
    """Emit the repeated `-c mcp_servers."excel-mcp".*` TOML overrides for *role*."""
    k, u = 'mcp_servers."excel-mcp".', json.dumps
    rows = (
        ("command", u(str(_venv_python(EXCEL_MCP_ROOT)))),
        ("args", u(["-m", "excel_mcp"])),
        ("cwd", u(str(EXCEL_MCP_ROOT))),
        ("startup_timeout_sec", "20"),
        ("tool_timeout_sec", "120"),
        ("enabled", "true"),
        ("enabled_tools", u(ROLE_TOOLS[role])),
    )
    return [p for n, v in rows for p in ("-c", f"{k}{n} = {v}")]


def build_codex_cmd(role: str, workspace: Path) -> list[str]:
    """Compose the `codex exec --json` argv for one agent invocation.

    Mirrors the proven-working CLI shape in `raw_api_baseline/runner.py`:
    `--ephemeral`, `--model <m>`, bare TOML enum values, OS-appropriate sandbox
    flag, and a trailing `-` so the prompt is delivered via stdin by the caller
    (see `BaseAgent._stream`). The excel-mcp overrides and `-C <workspace>`
    remain multi-agent-specific and are preserved.
    """
    if role not in ROLE_TOOLS:
        raise ValueError(f"unknown role {role!r}; expected one of {list(ROLE_TOOLS)}")
    cmd = [
        CODEX_CLI,
        "exec",
        "--json",
        "--ephemeral",
        "--skip-git-repo-check",
        "-C", str(workspace),
        "--model", MODEL,
        *_excel_mcp_overrides(role),
        # Bare (not double-quoted) enum identifiers; matches raw_api_baseline.
        "-c", f"model_reasoning_effort={REASONING_EFFORT}",
        "-c", f'service_tier="{SERVICE_TIER}"',
        "-c", f"features.fast_mode={str(FAST_MODE).lower()}",
        "-c", f"approval_policy={APPROVAL_POLICY}",
    ]
    if SANDBOX_MODE == "danger-full-access":
        cmd.append("--dangerously-bypass-approvals-and-sandbox")
    else:
        cmd.extend(["--sandbox", SANDBOX_MODE])
    cmd.append("-")  # prompt is piped to stdin by BaseAgent._stream
    return cmd
