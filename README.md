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

Add the server as a local `stdio` MCP process. Example project config:

```toml
[mcp_servers.excel]
command = "excel-mcp"
cwd = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp"
startup_timeout_sec = 20
tool_timeout_sec = 120
```

If you prefer to run from a virtual environment without installing the console script globally:

```toml
[mcp_servers.excel]
command = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp/.venv/bin/python"
args = ["-m", "excel_mcp"]
cwd = "/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp"
startup_timeout_sec = 20
tool_timeout_sec = 120
```

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
