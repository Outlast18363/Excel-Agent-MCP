---
name: excel-mcp
description: Use when tasks involve inspecting, editing, validating, or explaining complex Excel workbooks through the local `excel-mcp` MCP server. Best for live workbook sessions that need sheet discovery, targeted range reads, formula tracing, recalculation, or rendered screenshots.
---

# Excel MCP

## When to use

- Use this skill when the `excel-mcp` MCP server is installed and the task targets a live Excel workbook.
- Prefer this skill for complex `.xlsx` work that needs formula-aware edits, recalculation, screenshots, or dependency tracing.
- Use plain Python tooling instead when the task is limited to offline CSV/TSV analysis or simple workbook generation.

IMPORTANT: System and user instructions always take precedence.

## Quick start

1. Confirm the workbook path and whether the task is read-only or editable.
2. Call `open_workbook` first and keep the returned `workbook_id` for all later calls.
3. If the workbook or sheet is unfamiliar, call `get_sheet_state` before inspecting specific cells.
4. Use `get_range` to inspect the smallest range that gives enough context.
5. Use `set_range` for narrow, explicit edits instead of broad rewrites.
6. After input or formula changes, call `recalculate` before concluding the task.
7. Use `local_screenshot` when visual layout matters.
8. Use `trace_formula` when the task is about upstream logic, downstream impact, or formula lineage.

## Tool routing

- `open_workbook`: start or reuse the live workbook session. See [references/open_workbook.md](references/open_workbook.md)
- `get_sheet_state`: inspect used range, hidden structure, merged cells, and object counts. See [references/get_sheet_state.md](references/get_sheet_state.md)
- `get_range`: read targeted cell content, formulas, formats, styles, geometry, or merged info. See [references/get_range.md](references/get_range.md)
- `set_range`: write values, formulas, number formats, and lightweight styles. See [references/set_range.md](references/set_range.md)
- `recalculate`: trigger Excel recalculation and scan formula cells for errors. See [references/recalculate.md](references/recalculate.md)
- `local_screenshot`: export a rendered PNG of a range for review. See [references/local_screenshot.md](references/local_screenshot.md)
- `trace_formula`: trace precedents or dependents through the native workbook graph. See [references/trace_formula.md](references/trace_formula.md)

## Workflow and references

- Default workflow and decision heuristics: [references/workflow.md](references/workflow.md)
- Spreadsheet conventions, formatting rules, and formula guidance: [references/spreadsheet_guidelines.md](references/spreadsheet_guidelines.md)

## Working style

- Start from workbook state, not assumptions.
- Read narrowly before writing broadly.
- Prefer the smallest range that gives enough context.
- Recalculate after edits that could affect formulas.
- Use screenshots only when appearance matters.
- Keep changes intentional and reversible where possible.

## Notes

- This skill complements the MCP server; it does not replace MCP installation.
- If the local `excel-mcp` server is unavailable, tell the user and fall back to non-MCP spreadsheet tooling only when appropriate.
