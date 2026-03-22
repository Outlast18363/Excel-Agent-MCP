# Excel MCP

Lightweight Python MCP server for live Excel workbook interaction using `xlwings`.

This mcp aims to help advanced agents runing in codex or claude to smoothly explore and modify the most complex excel files: finiancial statement, multi-sheet mega-size files, Excel with sophisticated table & diagram structure, etc.

## MVP tools

- `open_workbook`
- `get_sheet_state`
- `get_range`
- `set_range`
- `recalculate`
- `local_screenshot`
- `trace_formula`

Every tool returns the same response envelope:

```json
{
  "status": "success",
  "data": {},
  "warnings": [],
  "errors": []
}
```

## Requirements

- Python 3.11+
- Microsoft Excel installed locally
- `xlwings` and the official Python `mcp` SDK
- Docker Desktop if you want to use `trace_formula`

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -e .
```

This install gives you both:

- the `excel-mcp` console command from `pyproject.toml`
- the module entrypoint `python -m excel_mcp`

## Run locally

The server is designed for `stdio` transport because that is the simplest and most compatible mode for local MCP clients.

```bash
excel-mcp
```

You can also run it as a module:

```bash
python3 -m excel_mcp
```

## TACO backend setup for `trace_formula`

`trace_formula` does **not** require the Excel add-in from `src/extern_service/tacolens/add-in/`.
For MCP usage, you only need the Java backend service from `src/extern_service/tacolens/`.

On macOS, after Docker Desktop is running, start the backend from the repository root with:

```bash
docker compose -f "src/extern_service/tacolens/docker-compose.yml" up --build
```

What this does:

- builds the bundled Java backend container
- starts the backend on `http://127.0.0.1:4567`
- keeps it running so `trace_formula` can call the TACO API

Notes:

- You do not need to run `npm install` or `npm run start` unless you want to use the TACO Excel add-in UI manually.
- The first Docker build may take a few minutes because Maven dependencies and the Java image need to be downloaded.
- Leave the Docker terminal running while using `trace_formula`.
- If you prefer to start it from inside the service folder instead, this is equivalent:

```bash
cd "src/extern_service/tacolens"
docker compose up --build
```

Quick checks:

- `docker compose -f "src/extern_service/tacolens/docker-compose.yml" ps`
- verify that the container exposes port `4567`
- if `trace_formula` says the backend is unavailable, check Docker logs from the same compose project

## Project layout

```text
excel_mcp/
├── .codex/
│   └── config.toml.example
├── pyproject.toml
├── README.md
├── src/
│   └── excel_mcp/
│       ├── __init__.py
│       ├── __main__.py
│       ├── server.py
│       ├── service.py
│       └── types.py
└── tests/
```

## Codex setup

Codex supports local `stdio` MCP servers either through the CLI or through `config.toml`. This project is a local `stdio` server and does not require extra environment variables for basic setup. The official Codex MCP docs describe both approaches and the supported config fields like `command`, `args`, `cwd`, `startup_timeout_sec`, and `tool_timeout_sec`.[https://developers.openai.com/codex/mcp](https://developers.openai.com/codex/mcp)

### Option 1: Add with the Codex CLI

From the repository root, after activating the virtualenv and installing the package:

```bash
source .venv/bin/activate
codex mcp add excel-mcp -- excel-mcp
```

If you prefer to avoid relying on the console script path, use the virtualenv interpreter directly:

```bash
codex mcp add excel-mcp -- \
  "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp/.venv/bin/python" -m excel_mcp
```

You can inspect configured servers with:

```bash
codex mcp --help
```

And in the Codex terminal UI, `/mcp` shows active MCP servers.[https://developers.openai.com/codex/mcp](https://developers.openai.com/codex/mcp)

### Option 2: Configure `config.toml`

Codex can read MCP configuration from either:

- `~/.codex/config.toml`
- a project-scoped `.codex/config.toml` in a trusted project

The repo includes a ready-to-copy example at `.codex/config.toml.example`.

#### Recommended project-scoped config

```toml
[mcp_servers.excel]
command = "excel-mcp"
cwd = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp"
startup_timeout_sec = 20
tool_timeout_sec = 120
```

#### Virtualenv-based config

```toml
[mcp_servers.excel]
command = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp/.venv/bin/python"
args = ["-m", "excel_mcp"]
cwd = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp"
startup_timeout_sec = 20
tool_timeout_sec = 120
```

Notes:

- `command` is required for a `stdio` MCP server.[https://developers.openai.com/codex/mcp](https://developers.openai.com/codex/mcp)
- `args` is optional and used here only for the virtualenv form.[https://developers.openai.com/codex/mcp](https://developers.openai.com/codex/mcp)
- `cwd` is recommended for this project so relative paths and local Docker commands behave predictably.[https://developers.openai.com/codex/mcp](https://developers.openai.com/codex/mcp)
- `startup_timeout_sec` defaults to `10` and `tool_timeout_sec` defaults to `60` in Codex; this repo recommends slightly higher values because opening Excel and tracing formulas can take longer.[https://developers.openai.com/codex/mcp](https://developers.openai.com/codex/mcp)
- If you want to temporarily disable the server or restrict tools, Codex also supports `enabled`, `required`, `enabled_tools`, and `disabled_tools` in the same config file.[https://developers.openai.com/codex/mcp](https://developers.openai.com/codex/mcp)

### Verify the Codex install

After adding the server, restart Codex if needed and confirm:

1. the server appears in `/mcp`
2. `open_workbook` and `get_range` are listed as available tools
3. `trace_formula` works after the TACO backend is running

### Files this repo already includes for Codex

This project already contains the runtime pieces Codex needs for a local `stdio` MCP server:

- `pyproject.toml`: declares the package and the `excel-mcp` console script
- `src/excel_mcp/__main__.py`: starts the server over stdio
- `src/excel_mcp/server.py`: defines the FastMCP tool surface
- `.codex/config.toml.example`: example local Codex configuration

So after installing the package and configuring Codex, you do not need any additional wrapper scripts just to run this MCP locally.

## Claude Code setup

Add the same server through Claude Code's MCP CLI:

```bash
claude mcp add --transport stdio excel-mcp -- excel-mcp
```

Or point Claude Code at the virtual environment interpreter:

```bash
claude mcp add --transport stdio excel-mcp -- \
  /Users/jz/Desktop/spreadsheet\ FINCH\ proj/excel_mcp/.venv/bin/python -m excel_mcp
```

## Notes about Excel behavior

- `open_workbook` keeps a process-local workbook registry keyed by `workbook_id`.
- `set_range` applies writes in a stable order and formulas overwrite values when both are supplied.
- `recalculate` scans only formula-bearing cells for errors.
- `local_screenshot` relies on xlwings' `Range.to_png()` feature which uses Excel's native export capability.
- `trace_formula` uses the bundled TACO-Lens backend to trace same-sheet precedents and dependents from the current worksheet formula graph.
- `trace_formula` expects the local TACO-Lens backend to be available at `http://127.0.0.1:4567/api/taco/patterns`.
- The server never writes logs to `stdout`; logs belong on `stderr` so MCP traffic stays clean.

## Tests

### Unit Tests
The included unit tests focus on helper logic and MCP wrapper behavior, which keeps them lightweight and runnable without a live Excel workbook:

```bash
cd "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp"
PYTHONPATH=src python3 -m unittest discover -s tests -p "test_*.py"
```

### End-to-End Tests
To verify absolute ground truth using your local Excel installation, you can run the live integration tests. This automatically generates a fresh Excel file on the fly, runs the actual MCP operations against it, and leaves output artifacts (including screenshots) behind so you can manually inspect accuracy.

To run the live E2E suite:
```bash
cd "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp"
PYTHONPATH=src python3 -m unittest tests.test_e2e
```

Notes:

- Run the commands from the repository root, not from `src/extern_service/tacolens/`.
- `python -m unittest` expects a module path like `tests.test_e2e`, not a slash path like `tests/test_e2e.py`.
- If you already installed the package with `python3 -m pip install -e .`, you can omit `PYTHONPATH=src`.
- Running `tests.test_e2e` will include the new `trace_formula` E2E cases too. They are in the same file.
- Some E2E tests may be skipped automatically if local prerequisites are missing, such as the Python `mcp` package, `xlwings`, or the TACO backend on port `4567`.

*Note: The script safely cleans up after itself. You'll find generated test workbooks and screenshot artifacts in `tests/test_output/` after it finishes running so you can inspect the results manually.*
