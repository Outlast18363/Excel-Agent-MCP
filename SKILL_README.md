# Excel MCP Skill Package

This repository includes an installable Codex skill package for the local `excel-mcp` MCP server.

## Package location

The skill package lives in `excel_mcp_skill/`.

## Contents

- `excel_mcp_skill/SKILL.md`: the main skill entry point that Codex discovers
- `excel_mcp_skill/agents/openai.yaml`: optional Codex metadata and MCP dependency declaration
- `excel_mcp_skill/references/`: progressive-disclosure docs for workflow and per-tool details

## Install

Copy or symlink `excel_mcp_skill/` to one of the locations Codex scans:

- `.agents/skills/excel-mcp/` for repo-local use
- `$HOME/.agents/skills/excel-mcp/` for user-global use

Then make sure the `excel-mcp` MCP server is installed in Codex. For repo-level install and configuration notes, see `skill_install.md`.
