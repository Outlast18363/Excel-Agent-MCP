# Excel MCP Skill Install

This file documents how to install the companion `excel-mcp` skill in Codex without putting human-oriented installation notes inside the agent-facing skill references.

## Codex skill package layout

The bundled skill package lives in `excel_mcp_skill/` and follows the Codex skill structure:

- required `SKILL.md`
- optional `agents/openai.yaml`
- optional `references/`

## Recommended install locations

- Repo-local install: `.agents/skills/excel-mcp/`
- User-local install: `$HOME/.agents/skills/excel-mcp/`

You can copy or symlink `excel_mcp_skill/` into either location.

## Manual install example

If you want this repository to expose the skill automatically, place the folder here:

```text
.agents/skills/excel-mcp/
```

The folder should contain at least:

```text
excel-mcp/
├── SKILL.md
├── agents/
│   └── openai.yaml
└── references/
    ├── workflow.md
    ├── spreadsheet_guidelines.md
    ├── open_workbook.md
    ├── get_sheet_state.md
    ├── get_range.md
    ├── set_range.md
    ├── recalculate.md
    ├── local_screenshot.md
    └── trace_formula.md
```

## MCP setup

This skill assumes the local `excel-mcp` server is already installed in Codex.

CLI setup example:

```bash
codex mcp add excel-mcp -- excel-mcp
```

If you want the virtualenv entrypoint instead:

```bash
codex mcp add excel-mcp -- \
  "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp/.venv/bin/python" -m excel_mcp
```

## `config.toml` examples

Codex supports both user-scoped `~/.codex/config.toml` and project-scoped `.codex/config.toml`.

Minimal stdio MCP server example:

```toml
[mcp_servers.excel-mcp]
command = "excel-mcp"
cwd = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp"
startup_timeout_sec = 20
tool_timeout_sec = 120
```

Virtualenv example:

```toml
[mcp_servers.excel-mcp]
command = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp/.venv/bin/python"
args = ["-m", "excel_mcp"]
cwd = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp"
startup_timeout_sec = 20
tool_timeout_sec = 120
```

To disable this skill without deleting it:

```toml
[[skills.config]]
path = "/absolute/path/to/excel-mcp/SKILL.md"
enabled = false
```

## Notes

- `agents/openai.yaml` declares the MCP dependency by server name, but it does not install the local server for you.
- `trace_formula` runs natively inside the Python MCP server and does not require a separate Taco or Docker backend.
- After adding or changing skills, restart Codex if the new skill does not appear immediately.
