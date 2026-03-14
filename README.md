# Excel MCP

Lightweight Python MCP server for live Excel workbook interaction using `xlwings`.

This mcp aims to help advanced agents runing in codex or claude to smoothly explore and modify the most complex excel files: finiancial statement, multi-sheet mega-size files, excel with sophesticated table & diagram structure, etc.

## MVP tools

- `open_workbook`
- `get_sheet_state`
- `get_range`
- `set_range`
- `recalculate`
- `local_screenshot`

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
- The server never writes logs to `stdout`; logs belong on `stderr` so MCP traffic stays clean.

## Tests

### Unit Tests
The included unit tests focus on helper logic and MCP wrapper behavior, which keeps them lightweight and runnable without a live Excel workbook:

```bash
python3 -m unittest discover -s tests -p "test_*.py"
```

### End-to-End Tests
To verify absolute ground truth using your local Excel installation, you can run the live integration tests. This automatically generates a fresh Excel file on the fly, runs the actual MCP operations against it, and leaves output artifacts (including screenshots) behind so you can manually inspect accuracy.

To run the live E2E suite:
```bash
python3 -m unittest tests/test_e2e.py
```

*Note: The script safely cleans up after itself. You'll find `test_workbook.xlsx` and `screenshot.png` in the `tests/test_output/` folder after it finishes running. These are left intentionally so you can visually verify the table structure was screenshotted perfectly.*
