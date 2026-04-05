---

## name: excel-mcp
description: Use when tasks involve inspecting, editing, validating, or explaining complex Excel workbooks through the local `excel-mcp` MCP server. Best for live workbook sessions that need sheet discovery, targeted range reads, formula tracing, recalculation, or rendered screenshots.

# Excel MCP

## When to use

- Inspect, edit, analyze, or visualize `.xlsx` workbooks via a live Excel session.
- Read or modify cell values, formulas, number formats, and styles in place.
- Recalculate formulas and audit errors before delivery.
- Capture rendered screenshots for visual review or layout verification.
- Trace formula precedents and dependents to understand workbook logic.
- Use plain Python tooling (`openpyxl`, `pandas`) instead when the task is limited to offline CSV/TSV analysis or simple workbook generation without Excel-native behavior.

IMPORTANT: System and user instructions always take precedence.

## Workflow

1. **Confirm task type** — create, edit, analyze, visualize, or explain workbook logic.
2. **Open the workbook** — call `open_workbook` to get the `workbook_id` needed by all later tools. Use `read_only=True` for investigative tasks.
3. **Orient at the sheet level** — call `get_sheet_state` to learn used range, hidden structure, merged cells, and formula density before inspecting specific cells.
4. **Inspect the target region** — call `get_range` with the smallest range that gives enough context. Choose `include_`* flags based on the task: `values` for content, `formulas` for logic, `styles`/`number_formats` for presentation, `geometry`/`merged_info` for layout. For ranges with cell number > 100, delegate the returned payload to a **sub-agent** to summarize key findings rather than processing the full result inline.
5. **Make narrow, explicit edits** — call `set_range` to write values, formulas, number formats, or styles. Prefer targeted edits over broad rewrites.
6. **Recalculate** — after any value or formula change, call `recalculate`. Match scope to blast radius: `range` for local edits, `sheet` for sheet-level changes, `workbook` for cross-sheet or uncertain impact. Inspect error summaries before concluding.
7. **Visual review** — call `local_screenshot` when layout or formatting matters. For large or full-sheet ranges where `local_screenshot` is impractical, use the LibreOffice approach instead (see Rendering below).
8. **Trace dependencies** — call `trace_formula` when you need to understand what feeds a cell (`precedents`) or what downstream cells depend on it (`dependents`).
9. **Save and clean up** — pass `save_after=True` on the final `set_range`, keep filenames stable, and delete intermediate files.

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


## Editing strategy

- Use `set_range` as the default edit path for values, formulas, number formats, and basic styles such as fill, font, alignment, and wrap. This keeps edits token-efficient and inside the active MCP workbook session.
- Do not use `set_range` for borders, merge or unmerge, insert or delete rows/columns, column width, row height, conditional formatting, data validation, comments or notes, copy/paste with formatting preserved, or AutoFit.
- If you need one of those unsupported operations, first call `close_workbook` with `save=True`, then perform the edit with `openpyxl`, then call `open_workbook` again before continuing with MCP-based inspection, recalculation, or screenshots.
- Never `import xlwings` in a standalone script while an MCP workbook session is open. That creates a second COM connection to the same workbook and can cause lock conflicts, stale handles, or divergent session state.

## Rendering and visual checks

- **Preferred**: use `local_screenshot` for targeted range review via the MCP server.
- **Large or full-sheet ranges**: if the range is too big for `local_screenshot`, use LibreOffice + Poppler:
  ```
  soffice --headless --convert-to pdf --outdir $OUTDIR $INPUT_XLSX
  pdftoppm -png $OUTDIR/$BASENAME.pdf $OUTDIR/$BASENAME
  ```
- If no rendering tools are available, tell the user that layout should be reviewed locally.
- Review rendered output for layout, formula results, clipping, inconsistent styles, and spilled text.

## Temp and output conventions

- Use `tmp/spreadsheets/` for intermediate files; delete them when done.
- Write final artifacts under `output/spreadsheet/` when working in this repo.
- Keep filenames stable and descriptive.

## Fallback tooling (non-MCP)

- `openpyxl` for creating/editing `.xlsx` files and preserving formatting offline.
- `pandas` for analysis and CSV/TSV workflows; write results back to `.xlsx` or `.csv`.
- `openpyxl.chart` for native Excel charts when MCP charting is not available.

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

## Citation requirements

- Cite sources inside the spreadsheet using plain-text URLs.
- For financial models, cite model inputs in cell comments.
- For tabular data sourced externally, add a source column when each row represents a separate item.

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

## Color conventions (if no style guidance)

- Blue: user input
- Black: formulas and derived values
- Green: linked or imported values
- Gray: static constants
- Orange: review or caution
- Light red: error or flag
- Purple: control or logic
- Teal: visualization anchors and KPI highlights

## Finance-specific requirements

- Format zeros as `-`.
- Negative numbers should be red and in parentheses.
- Format multiples as `5.2x`.
- Always specify units in headers (e.g. `Revenue ($mm)`).
- Cite sources for all raw inputs in cell comments.
- For new financial models with no user-specified style, use blue text for hardcoded inputs, black for formulas, green for internal workbook links, red for external links, and yellow fill for key assumptions that need attention.

## Investment banking layouts

If the spreadsheet is an IB-style model (LBO, DCF, 3-statement, valuation):

- Totals should sum the range directly above.
- Hide gridlines and use horizontal borders above totals across relevant columns.
- Section headers should be merged cells with dark fill and white text.
- Column labels for numeric data should be right-aligned; row labels should be left-aligned.
- Indent submetrics under their parent line items.

## Notes

- If you rely on internal spreadsheet tooling, do not expose internal code or private APIs in user-facing explanations or code samples.
- `openpyxl` does not evaluate formulas; preserve formulas and use MCP `recalculate` when available.

