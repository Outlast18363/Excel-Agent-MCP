"""Codex CLI command builder shared by all two-agent roles.

Custom providers use Codex's supported per-run config overrides. Put local
provider credentials in ``two_agent_framework/local_provider.json``; the API key
is injected only into the ``codex exec`` child process environment.

The ``TWOWORK_CODEX_*`` environment variables remain available as a fallback
when no local provider config is present. ``NODE_BINARY`` optionally sets the
Node path when using the ``node .../codex.js`` launch on Windows.
"""

import json
import os
import shutil
import sys
from pathlib import Path


def _excel_mcp_root() -> Path:
    """Project root that contains ``two_agent_framework`` (honors ``EXCEL_MCP_ROOT``)."""
    return Path(os.environ.get("EXCEL_MCP_ROOT") or Path(__file__).resolve().parents[1])


def codex_launch_prefix() -> list[str]:
    """Argv fragment before ``exec``: ``[codex]`` or Windows ``[node, codex.js]``.

    Running ``codex.cmd`` directly under ``subprocess`` has been observed to fail with
    ``os error 2`` while ``node codex.js`` works for the same arguments.
    """
    explicit = (os.environ.get("CODEX_CLI") or os.environ.get("CODEX_BINARY") or "").strip()
    if explicit:
        return [explicit]
    which = shutil.which("codex")
    if os.name == "nt" and which:
        p = Path(which)
        if p.suffix.lower() in (".cmd", ".bat") and p.name.lower().startswith("codex"):
            js = p.parent / "node_modules" / "@openai" / "codex" / "bin" / "codex.js"
            if js.is_file():
                node = (os.environ.get("NODE_BINARY") or "").strip()
                candidates = []
                if node:
                    candidates.append(Path(node))
                candidates.append(Path(r"C:\Program Files\nodejs\node.exe"))
                wn = shutil.which("node")
                if wn:
                    candidates.append(Path(wn))
                for node_exe in candidates:
                    if node_exe.is_file():
                        return [str(node_exe.resolve()), str(js.resolve())]
    if which:
        return [which]
    return ["codex"]


EXCEL_MCP_ROOT = _excel_mcp_root()
LOCAL_PROVIDER_CONFIG = Path(__file__).with_name("local_provider.json")
APPROVAL_POLICY = "never"
# On Windows, codex's `workspace-write` sandbox has been observed to cause native
# sandbox access failures (documented in raw_api_baseline/README.md). Match the
# raw_api_baseline default there: `danger-full-access` paired with the
# `--dangerously-bypass-approvals-and-sandbox` flag below. macOS/Linux keep the
# original `workspace-write` behavior the framework was written for.
SANDBOX_MODE = "danger-full-access" if os.name == "nt" else "workspace-write"

# Per-role excel-mcp tool allowlists. Worker does discovery plus implementation;
# Evaluator verifies from snapshots/final outputs; Distiller is pure text.
_WORKER_TOOLS = [
    "open_workbook",
    # "get_sheet_state",
    "local_screenshot",
    "close_workbook",
    # "search_cell",
    "web_search",
    "get_range",
]
_EVALUATOR_TOOLS = [
    "open_workbook",
    # "get_sheet_state",
    "local_screenshot",
    "close_workbook",
    # "search_cell",
    "get_range",
]
ROLE_TOOLS: dict[str, list[str]] = {
    "Worker": _WORKER_TOOLS,
    "Evaluator": _EVALUATOR_TOOLS,
    "Distiller": [],
}


def _venv_python(root: Path) -> Path:
    """Return the virtualenv Python path for the current platform."""
    if sys.platform.startswith("win"):
        return root / ".venv" / "Scripts" / "python.exe"
    return root / ".venv" / "bin" / "python"


def twowork_subprocess_env() -> dict[str, str]:
    """Return the environment for Codex, with local provider secrets scoped to the child."""
    env = os.environ.copy()
    config = _local_provider_config()
    env_key = config.get("env_key", "")
    api_key = config.get("api_key", "")
    if env_key and api_key:
        env[env_key] = api_key
    return env


def _toml_scalar_str(value: str) -> str:
    """JSON/TOML scalar string for `-c foo=...`."""
    return json.dumps(value)


def _env_truthy(key: str) -> bool:
    return (os.environ.get(key) or "").strip().lower() in {"1", "true", "yes", "on"}


def _local_provider_config() -> dict[str, str]:
    """Load optional local custom-provider config."""
    if _env_truthy("TWOWORK_CODEX_DISABLE_LOCAL_PROVIDER"):
        return {}
    if not LOCAL_PROVIDER_CONFIG.is_file():
        return {}
    try:
        payload = json.loads(LOCAL_PROVIDER_CONFIG.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"invalid JSON in {LOCAL_PROVIDER_CONFIG}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError(f"{LOCAL_PROVIDER_CONFIG} must contain a JSON object")
    return {
        str(key): str(value).strip()
        for key, value in payload.items()
        if value is not None and str(value).strip()
    }


def _config_value(config: dict[str, str], key: str, env_key: str) -> str:
    return config.get(key) or (os.environ.get(env_key) or "").strip()


def _provider_wire_api(config: dict[str, str], provider: str) -> str:
    """Return the Codex wire API for a custom provider.

    Codex CLI 0.118+ rejects the legacy Chat Completions wire API for custom
    providers. OpenRouter's Responses endpoint is still a beta compatibility
    layer, so use direct OpenAI/ChatGPT auth for production Worker runs that
    require robust shell/file tool continuation.
    """
    wire_api = _config_value(config, "wire_api", "TWOWORK_CODEX_PROVIDER_WIRE_API")
    return wire_api or "responses"


def _codex_provider_overrides() -> list[str]:
    """Return optional Codex `-c` overrides for a custom model provider."""
    config = _local_provider_config()
    provider = _config_value(config, "provider_id", "TWOWORK_CODEX_PROVIDER_ID")
    if not provider:
        return []

    base_url = _config_value(config, "base_url", "TWOWORK_CODEX_PROVIDER_BASE_URL").rstrip("/")
    env_key = _config_value(config, "env_key", "TWOWORK_CODEX_PROVIDER_ENV_KEY")
    name = _config_value(config, "provider_name", "TWOWORK_CODEX_PROVIDER_NAME") or provider
    wire_api = _provider_wire_api(config, provider)

    rows = [
        ("model_provider", _toml_scalar_str(provider)),
        (f"model_providers.{provider}.name", _toml_scalar_str(name)),
    ]
    if base_url:
        rows.append((f"model_providers.{provider}.base_url", _toml_scalar_str(base_url)))
    if env_key:
        rows.append((f"model_providers.{provider}.env_key", _toml_scalar_str(env_key)))
    if wire_api:
        rows.append((f"model_providers.{provider}.wire_api", _toml_scalar_str(wire_api)))
    return [part for key, value in rows for part in ("-c", f"{key}={value}")]


def effective_codex_model() -> str | None:
    """Model id for ``--model``; honors ``TWOWORK_CODEX_MODEL``."""
    override = _config_value(_local_provider_config(), "model", "TWOWORK_CODEX_MODEL")
    if override:
        return override
    return None


def _reasoning_overrides() -> list[str]:
    effort = _config_value(_local_provider_config(), "reasoning_effort", "TWOWORK_CODEX_REASONING_EFFORT")
    if not effort:
        return []
    return ["-c", f"model_reasoning_effort={effort}"]


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
    parts = [p for n, v in rows for p in ("-c", f"{k}{n} = {v}")]
    return parts


def build_codex_cmd(role: str, workspace: Path, *, with_excel_mcp: bool = True) -> list[str]:
    """Compose the `codex exec --json` argv for one agent invocation.

    The excel-mcp overrides and `-C <workspace>` are framework-specific. Custom
    providers are configured only through Codex CLI's supported per-run `-c`
    runtime overrides.

    Set ``with_excel_mcp=False`` for lightweight probes that only need the same
    Codex launch shape without starting excel-mcp.
    """
    if role not in ROLE_TOOLS:
        raise ValueError(f"unknown role {role!r}; expected one of {list(ROLE_TOOLS)}")
    model = effective_codex_model()

    mcp_fragments = _excel_mcp_overrides(role) if with_excel_mcp else []
    model_fragments = ["--model", model] if model else []

    cmd = [
        *codex_launch_prefix(),
        "exec",
        "--json",
        "--ephemeral",
        "--skip-git-repo-check",
        "-C",
        str(workspace),
        *model_fragments,
        *_codex_provider_overrides(),
        *_reasoning_overrides(),
        *mcp_fragments,
        "-c",
        f"approval_policy={APPROVAL_POLICY}",
    ]
    if SANDBOX_MODE == "danger-full-access":
        cmd.append("--dangerously-bypass-approvals-and-sandbox")
    else:
        cmd.extend(["--sandbox", SANDBOX_MODE])
    cmd.append("-")  # prompt is piped to stdin by BaseAgent._stream
    return cmd
