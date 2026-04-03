# local_screenshot

## Purpose

Captures the rendered on-screen appearance of a cell range in an open Excel workbook as a PNG image.

## When to use

Use when you need a visual snapshot of how Excel displays a range (formatting, borders, merged cells) rather than raw cell values or formulas. This is range-scoped output, not a full-monitor screenshot.

## Parameters

- **workbook_id** (`str`, required) — ID of the workbook opened with `open_workbook`.
- **sheet** (`str`, required) — Sheet name.
- **range** (`str`, required) — A1-style address of the range to render.
- **output_path** (`str | None`, default `None`) — Filesystem path for the PNG. If omitted, a temporary file is created (prefix `excel-mcp-` in the system temp directory). Parent directories are created when needed.
- **return_base64** (`bool`, default `False`) — When `True`, the response `data` includes a base64-encoded PNG payload.

## Response `data` fields

- **sheet** (`str`) — Sheet name that was captured.
- **range** (`str`) — The A1 address of the captured range.
- **image_path** (`str`) — Path to the written PNG file.
- **base64** (`str`) — Base64-encoded PNG bytes; present only when `return_base64=True`.

## Notes

- Implementation uses xlwings `Range.to_png(...)`. The raw export is RGBA with a transparent background for unfilled cells; the service composites onto white and converts to RGB so every pixel is fully opaque.
- If `output_path` is omitted, the image is written to a temporary PNG path returned in `image_path`.
- This captures only the specified Excel range, not the entire monitor or window.

## Example

```python
server.local_screenshot(workbook_id, "Data", "A1:D10", output_path="/tmp/data_preview.png")
```
