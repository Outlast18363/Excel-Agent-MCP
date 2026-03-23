# Spreadsheet Guidelines

Use these rules when the task is spreadsheet-heavy but not limited to one specific MCP tool call.

## File Type And Goal

1. Confirm whether the task is create, edit, analyze, visualize, or explain workbook logic.
2. Prefer the live MCP workflow for complex `.xlsx` tasks that need Excel-native behavior.
3. Prefer `openpyxl` for offline `.xlsx` editing and formatting.
4. Prefer `pandas` for analysis and CSV/TSV workflows.

## Temp And Output Conventions

- Use `tmp/spreadsheets/` for intermediate files; delete them when done.
- Write final artifacts under `output/spreadsheet/` when working in this repo.
- Keep filenames stable and descriptive.

## Primary Tooling

- Use `openpyxl` for creating or editing `.xlsx` files while preserving formatting.
- Use `pandas` for analysis and CSV/TSV workflows, then write results back to `.xlsx` or `.csv`.
- Use `openpyxl.chart` for native Excel charts when needed.
- If the local Excel MCP server is available, prefer it for recalculation, screenshots, and dependency tracing.

## Recalculation And Visual Review

- Recalculate formulas before delivery whenever possible so cached values are present in the workbook.
- Render each relevant sheet or range for visual review when rendering tooling is available.
- `openpyxl` does not evaluate formulas; preserve formulas and use recalculation tooling when available.
- If you rely on internal spreadsheet tooling, do not expose internal code or private APIs in user-facing explanations or code samples.

## Rendering And Visual Checks

- If repo-local rendering is available, prefer `local_screenshot` for targeted sheet or range review.
- If LibreOffice (`soffice`) and Poppler (`pdftoppm`) are available, render sheets for visual review:
  - `soffice --headless --convert-to pdf --outdir $OUTDIR $INPUT_XLSX`
  - `pdftoppm -png $OUTDIR/$BASENAME.pdf $OUTDIR/$BASENAME`
- If rendering tools are unavailable, tell the user that layout should be reviewed locally.
- Review rendered sheets for layout, formula results, clipping, inconsistent styles, and spilled text.

## Formula Requirements

- Use formulas for derived values rather than hardcoding results.
- Do not use dynamic array functions like `FILTER`, `XLOOKUP`, `SORT`, or `SEQUENCE`.
- Keep formulas simple and legible; use helper cells for complex logic.
- Avoid volatile functions like `INDIRECT` and `OFFSET` unless required.
- Prefer cell references over magic numbers.
- Use absolute (`$B$4`) or relative (`B4`) references carefully so copied formulas behave correctly.
- If you need literal text that starts with `=`, prefix it with a single quote.
- Guard against `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, and `#NAME?` errors.
- Check for off-by-one mistakes, circular references, and incorrect ranges.

## Citation Requirements

- Cite sources inside the spreadsheet using plain-text URLs.
- For financial models, cite model inputs in cell comments.
- For tabular data sourced externally, add a source column when each row represents a separate item.

## Formatting Requirements For Existing Spreadsheets

- Render and inspect a provided spreadsheet before modifying it when possible.
- Preserve existing formatting and style exactly.
- Match styles for any newly filled cells that were previously blank.
- Never overwrite established formatting unless the user explicitly asks for a redesign.

## Formatting Requirements For New Or Unstyled Spreadsheets

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

## Color Conventions If No Style Guidance Exists

- Blue: user input
- Black: formulas and derived values
- Green: linked or imported values
- Gray: static constants
- Orange: review or caution
- Light red: error or flag
- Purple: control or logic
- Teal: visualization anchors and KPI highlights

## Finance-Specific Requirements

- Format zeros as `-`.
- Negative numbers should be red and in parentheses.
- Format multiples as `5.2x`.
- Always specify units in headers.
- Cite sources for all raw inputs in cell comments.
- For new financial models with no user-specified style, use blue text for hardcoded inputs, black for formulas, green for internal workbook links, red for external links, and yellow fill for key assumptions that need attention.

## Investment Banking Layouts

If the spreadsheet is an IB-style model such as LBO, DCF, 3-statement, or valuation:

- Totals should sum the range directly above.
- Hide gridlines and use horizontal borders above totals across relevant columns.
- Section headers should be merged cells with dark fill and white text.
- Column labels for numeric data should be right-aligned; row labels should be left-aligned.
- Indent submetrics under their parent line items.
