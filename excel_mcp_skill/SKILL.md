---
name: "excel-mcp"
description: "Use when tasks involve inspecting, editing, validating, or explaining complex Excel workbooks through the local `excel-mcp` MCP server. Best for live workbook sessions that need sheet discovery, targeted range reads, formula tracing, recalculation, or rendered screenshots."
---

# Excel MCP

## When to use

- Inspect, edit, analyze, or visualize `.xlsx` workbooks via a live Excel session.
- Read or modify cell values, formulas, number formats, and styles in place.
- Recalculate formulas and audit errors before delivery.
- Capture rendered screenshots for visual review or layout verification.
- Trace formula precedents and dependents to understand workbook logic.

IMPORTANT: System and user instructions always take precedence.



## Workflow

1. **Confirm task type** — create, edit, analyze, visualize, or explain workbook logic.
2. **Open the workbook** — call `open_workbook` to get the `workbook_id` needed by all later tools. Use `read_only=True` for investigative tasks.
3. **Orient at the sheet level** — call `get_sheet_state` to learn used range, hidden structure, merged cells, and formula density before inspecting specific cells.
4. **Devise a brief edit plan for the task.** For example, the plan should include the rough phases for the edit, or the information and formula you need for the edit. The point is to guide yourself so you don't overread the Excel.
5. **Make intentional and well-formatted edits.**
6. You may use functions like 'get_range', 'local_screenshot', and 'trace_formula' to guide your edits.



## Important Practices

1. **Inspect the target region** — call a **sub-agent** to call `get_range` with the smallest range that gives enough context, and return to you the needed summary of the key findings of the range. Choose `include_`* flags based on the task: `values` for content, `formulas` for logic, `styles`/`number_formats` for presentation, `geometry`/`merged_info` for layout. DON'T use get_range on ranges with more than 200 cells (estimate before you call subagent), use local_screenshot instead.

2. **Make narrow, explicit edits (STRICTLY FOLLOW the formatting requirement of the task)** 

3. **Recalculate** — after any value or formula change, call `recalculate`. Match scope to blast radius: `range` for local edits, `sheet` for sheet-level changes, `workbook` for cross-sheet or uncertain impact. Inspect error summaries before concluding.

4. **Visual review** — call `local_screenshot` when layout or formatting matters or you want to query a large range ( >300 cells)

5. **Trace dependencies** — call `trace_formula` when you need to understand what feeds a cell (`precedents`) or what downstream cells depend on it (`dependents`).

6. **Save and clean up** — pass `save_after=True` on the final `set_range`, keep filenames stable, and delete intermediate files.

7. Use `set_range` as the default edit path for values, formulas, number formats, and basic styles such as fill, font, alignment, and wrap.  

8. If you need an editing feature unsupported by 'set_range', first call `close_workbook` with `save=True`, then perform the edit with `openpyxl`, then call `open_workbook` again before continuing with MCP-based inspection, recalculation, or screenshots.

9. Never `import xlwings` in a standalone script while an MCP workbook session is open. That creates a second COM connection to the same workbook and can cause lock conflicts, stale handles, or divergent session state.

   

## MCP tool routing


| Tool               | Purpose                                                                                  | Reference                                             |
| ------------------ | ---------------------------------------------------------------------------------------- | ----------------------------------------------------- |
| `open_workbook`    | Start or reuse a live workbook session                                                   | [open_workbook.md](references/open_workbook.md)       |
| `get_sheet_state`  | Sheet metadata: used range, hidden rows/cols, merged cells, formula/object counts        | [get_sheet_state.md](references/get_sheet_state.md)   |
| `get_range`        | Read cell values, formulas, formats, styles, geometry, or merge info for a bounded range | [get_range.md](references/get_range.md)               |
| `set_range`        | Write values, formulas, number formats, and lightweight styles to a range                | [set_range.md](references/set_range.md)               |
| `recalculate`      | Trigger Excel recalculation and scan for formula errors                                  | [recalculate.md](references/recalculate.md)           |
| `local_screenshot` | Render a range as a PNG for visual review                                                | [local_screenshot.md](references/local_screenshot.md) |
| `trace_formula`    | Trace precedents or dependents through the native workbook graph                         | [trace_formula.md](references/trace_formula.md)       |


## 

## Temp and output conventions

- Use `tmp/spreadsheets/` for intermediate files; delete them when done.
- Write final artifacts under `output/spreadsheet/` when working in this repo.
- Keep filenames stable and descriptive.

## Formula requirements

- Use formulas for derived values rather than hardcoding results.
- Do not use dynamic array functions like `FILTER`, `XLOOKUP`, `SORT`, or `SEQUENCE`.
- Keep formulas simple and legible; use helper cells for complex logic.
- Avoid volatile functions like `INDIRECT` and `OFFSET` unless required.
- Prefer cell references over magic numbers (e.g. `=H6*(1+$B$3)` instead of `=H6*1.04`).
- Use absolute (`$B$4`) or relative (`B4`) references carefully so copied formulas behave correctly.
- If you need literal text that starts with `=`, prefix it with a single quote.
- Guard against `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, and `#NAME?` errors.
- Check for off-by-one mistakes, circular references, and incorrect ranges.


## Formatting requirements (existing spreadsheets)

- Render and inspect a provided spreadsheet before modifying it when possible.
- Preserve existing formatting and style exactly.
- Match styles for any newly filled cells that were previously blank.
- Never overwrite established formatting unless the user explicitly asks for a redesign.

## Formatting requirements (new or unstyled spreadsheets)

- Use appropriate number and date formats.
- Dates should render as dates, not plain numbers.
- Percentages should usually default to one decimal place unless the data calls for something else.
- Currencies should use the appropriate currency format.
- Headers should be visually distinct from raw inputs and derived cells.
- Use fill colors, borders, spacing, and merged cells sparingly and intentionally.
- Set row heights and column widths so content is readable without excessive whitespace.
- Do not apply borders around every filled cell.
- Group related calculations and make totals simple sums of the cells above them.
- Add whitespace to separate sections.
- Ensure text does not spill into adjacent cells.
- Avoid unsupported spreadsheet data-table features such as `=TABLE`.


## Notes

- `openpyxl` does not evaluate formulas; preserve formulas and use MCP `recalculate` when available.

