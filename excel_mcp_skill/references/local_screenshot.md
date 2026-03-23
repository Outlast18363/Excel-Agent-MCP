# `local_screenshot`

## What It Does

Exports a PNG image of a rendered Excel range.

## When To Use It

Use this when the visual result matters, such as checking layout, alignment, report presentation, or final formatting.

## Parameters

- `workbook_id: str`
- `sheet: str`
- `range: str`
- `output_path: str | None = None`
- `return_base64: bool = False`

## Returns In `data`

- `sheet`
- `range`
- `image_path`
- `base64`

## Notes

- This is not a screen capture of the whole monitor.
- The implementation uses Excel/xlwings range export through `Range.to_png(...)`.
- If `output_path` is omitted, the service creates a temporary PNG file path.
- `base64` is returned only when `return_base64=True`.

## Example

```python
local_screenshot(
    workbook_id="wb_001",
    sheet="Dashboard",
    range="A1:J25",
    output_path="/tmp/dashboard.png",
)
```
