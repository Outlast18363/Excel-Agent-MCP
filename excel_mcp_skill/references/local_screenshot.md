# local_screenshot

## Purpose

Captures the rendered on-screen appearance of a cell range in an open Excel workbook as a PNG image.

## When to use

Use when you need a visual snapshot of how Excel displays a range (formatting, borders, merged cells) rather than raw cell values or formulas. This is range-scoped output, not a full-monitor screenshot.

## Parameters

- **workbook_id** (`str`, required) — ID of the workbook opened with `open_workbook`.
- **sheet** (`str`, required) — Sheet name.
- **range** (`str`, required) — A1-style address of the range to render.
- **output_path** (`str | None`, default `None`) — Filesystem path for the PNG. If omitted, the server writes a stable file under `output/spreadsheet/screenshots/`. Parent directories are created when needed.

## Response `data` fields

- **sheet** (`str`) — Sheet name that was captured.
- **range** (`str`) — The A1 address of the captured range.
- **image_path** (`str`) — Path to the written PNG file.

## Notes

- If `output_path` is omitted, use the returned `image_path` under `output/spreadsheet/screenshots/` rather than expecting inline image bytes.
- This captures only the specified Excel range, not the entire monitor or window.

## Example

```python
server.local_screenshot(
    workbook_id,
    "Data",
    "A1:D10",
    output_path="output/spreadsheet/screenshots/data_preview.png",
)
```
