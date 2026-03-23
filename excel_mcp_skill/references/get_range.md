# `get_range`

## What It Does

Reads an explicit A1 range and returns:

- `matrix`: a row-major 2D layout
- `cells`: a flat list of cell payloads

Depending on flags, the payload can include values, formulas, number formats, styles, geometry, hidden flags, and merged-cell information.

## When To Use It

Use this when you already know the target range and need actual workbook contents rather than sheet-level structure.

## Parameters

- `workbook_id: str`
- `sheet: str`
- `range: str`
- `include_values: bool = True`
- `include_formulas: bool = False`
- `include_number_formats: bool = False`
- `include_styles: bool = False`
- `include_geometry: bool = False`
- `include_hidden_flags: bool = False`
- `include_merged_info: bool = False`

## Returns In `data`

- `sheet`
- `range`
- `matrix`
- `cells`

Each cell always includes:

- `address`
- `row`
- `column`

Optional cell fields include:

- `value`
- `formula`
- `number_format`
- `style`
- `geometry`
- `row_hidden`
- `column_hidden`
- `is_merged`
- `merged_range`

If `include_geometry=True`, the top-level payload also includes range geometry.

## Notes

- For non-formula cells, `formula` is returned as `null` when formulas are requested.
- `style` is shallow and meant for inspection, not full Excel style round-tripping.
- Use `matrix` when position matters and `cells` when filtering or scanning is easier.

## Example

```python
get_range(
    workbook_id="wb_001",
    sheet="Income Statement",
    range="B4:E12",
    include_values=True,
    include_formulas=True,
    include_number_formats=True,
    include_styles=True,
)
```
