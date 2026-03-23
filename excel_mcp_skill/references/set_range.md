# `set_range`

## What It Does

Writes values, formulas, number formats, and lightweight styles into an explicit A1 range.

It can also clear contents first, clear formatting first, and optionally save the workbook after the write.

## When To Use It

Use this for live workbook edits such as entering inputs, replacing formulas, clearing a block, or applying basic formatting.

## Parameters

- `workbook_id: str`
- `sheet: str`
- `range: str`
- `values: Any = None`
- `formulas: Any = None`
- `number_format: Any = None`
- `style: dict[str, Any] | None = None`
- `clear_contents: bool = False`
- `clear_formats: bool = False`
- `save_after: bool = False`

## Supported `style` Keys

- `fill_color`
- `font_name`
- `font_size`
- `font_bold`
- `font_italic`
- `font_color`
- `horizontal_alignment`
- `vertical_alignment`
- `wrap_text`

## Returns In `data`

- `sheet`
- `range`
- `updated_values`
- `updated_formulas`
- `updated_style`
- `saved`

## Notes

- Multi-cell `values`, `formulas`, and matrix `number_format` payloads must match the target range shape exactly.
- A scalar is accepted for a single-cell target and normalized to a `1 x 1` matrix.
- If both `values` and `formulas` are provided, formulas are written after values, so formulas win.
- This tool marks the cached trace graph as dirty, so a later trace may rebuild even if `refresh_graph=False`.

## Example

```python
set_range(
    workbook_id="wb_001",
    sheet="Report",
    range="B2:D2",
    style={
        "font_bold": True,
        "fill_color": "#D9EAF7",
        "horizontal_alignment": "center",
    },
    number_format="$#,##0.00",
)
```
