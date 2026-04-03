# get_range

## Purpose

Returns a row-major `matrix` and a matching flat `cells` list for an explicit A1 range, with optional values, formulas, formats, styles, geometry, hidden-row/column flags, and merged-cell metadata.

## When to use

Use this when you already know the sheet and range and need cell-level content from the live workbook, rather than whole-sheet structure from `get_sheet_state`.

## Parameters

- `workbook_id` (`str`, required) — Session handle from `open_workbook`.
- `sheet` (`str`, required) — Sheet name containing the range.
- `range` (`str`, required) — A1-style address of the region to read (e.g. `B4:E12`).
- `include_values` (`bool`, default `True`) — Include each cell’s computed `value`.
- `include_formulas` (`bool`, default `False`) — Include `formula` strings (`=` prefix) where applicable.
- `include_number_formats` (`bool`, default `False`) — Include Excel `number_format` strings.
- `include_styles` (`bool`, default `False`) — Include shallow per-cell `style` objects.
- `include_geometry` (`bool`, default `False`) — Include per-cell `geometry` and top-level range `geometry`.
- `include_hidden_flags` (`bool`, default `False`) — Include `row_hidden` and `column_hidden` on each cell.
- `include_merged_info` (`bool`, default `False`) — Include `is_merged` and, when merged, `merged_range`.

## Response `data` fields

Always present:

- `sheet` (`str`) — Sheet name.
- `range` (`str`) — Resolved A1 address of the read range (may differ from the input if Excel normalizes it).
- `matrix` (`list[list[cell]]`) — Row-major 2D grid of cell objects.
- `cells` (`list[cell]`) — Same cell objects as `matrix`, in row-major order (each entry is the identical object as the corresponding `matrix[row][col]`).

Each `cell` always includes:

- `address` (`str`) — e.g. `B2`.
- `row` (`int`) — 1-based row index.
- `column` (`int`) — 1-based column index.

Optional fields (only when the matching `include_*` flag is true):

- `value` — Computed cell value (`include_values`).
- `formula` (`str` | `null`) — Formula string starting with `=`, or `null` for non-formula cells (`include_formulas`).
- `number_format` (`str` | `null`) — Excel format code (`include_number_formats`).
- `style` (`object`) — Shallow style snapshot (`include_styles`); keys:
  - `font_name` (`str` | `null`)
  - `font_size` (`float` | `null`)
  - `font_bold` (`bool` | `null`)
  - `font_italic` (`bool` | `null`)
  - `font_color` (normalized color value | `null`)
  - `horizontal_alignment` (`str` | `null`)
  - `vertical_alignment` (`str` | `null`)
  - `wrap_text` (`bool` | `null`)
  - `fill_color` (normalized color value | `null`)
- `geometry` (`object`) — Per-cell layout (`include_geometry`): `left`, `top`, `width`, `height` (floats, points).
- `row_hidden` (`bool`) — Whether the cell’s row is hidden (`include_hidden_flags`).
- `column_hidden` (`bool`) — Whether the cell’s column is hidden (`include_hidden_flags`).
- `is_merged` (`bool`) — Whether the cell participates in a merge (`include_merged_info`).
- `merged_range` (`str`) — A1 address of the merge area; present only when `is_merged` is true (`include_merged_info`).

When `include_geometry` is true, `data` also includes:

- `geometry` (`object`) — Range-level box: `left`, `top`, `width`, `height` (floats, points).

## Notes

- `formula` is `null` for cells that are not formulas, not an empty string.
- `style` is shallow and intended for inspection, not full Excel style round-tripping.
- Prefer `matrix` when row/column position matters; prefer `cells` when filtering, mapping, or scanning without nested indexing.

## Example

```python
get_range(
    workbook_id=workbook_id,
    sheet="Data",
    range="A1:B2",
    include_values=True,
    include_formulas=True,
    include_styles=True,
)
```
