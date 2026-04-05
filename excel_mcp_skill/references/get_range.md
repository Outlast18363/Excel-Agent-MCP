# get_range

## Purpose

Returns a compact payload for an explicit A1 range using dense range-aligned arrays instead of repeated per-cell objects.

## When to use

Use this when you already know the sheet and range and need live workbook content from a bounded region, especially for medium or large reads where duplicated per-cell metadata would be wasteful.

## Parameters

- `workbook_id` (`str`, required) — Session handle from `open_workbook`.
- `sheet` (`str`, required) — Sheet name containing the range.
- `range` (`str`, required) — A1-style address of the region to read (e.g. `B4:E12`).
- `include_values` (`bool`, default `True`) — Include each cell’s computed `value`.
- `include_formulas` (`bool`, default `False`) — Include a `formulas` matrix containing formula strings (`=` prefix) or `null` for non-formula cells.
- `include_number_formats` (`bool`, default `False`) — Include a `number_formats` matrix aligned to the requested range.
- `include_styles` (`bool`, default `False`) — Include a deduplicated `style_table` plus per-cell `style_ids`.
- `include_geometry` (`bool`, default `False`) — Include top-level range `geometry`.
- `include_hidden_flags` (`bool`, default `False`) — Include top-level `row_hidden` and `column_hidden` arrays aligned to the requested range.
- `include_merged_info` (`bool`, default `False`) — Include deduplicated `merged_ranges` once per response.

## Response `data` fields

Always present:

- `sheet` (`str`) — Sheet name.
- `range` (`str`) — Resolved A1 address of the read range (may differ from the input if Excel normalizes it).
- `rows` (`int`) — Number of rows in the requested range.
- `columns` (`int`) — Number of columns in the requested range.

Optional fields (only when the matching `include_*` flag is true):

- `values` (`list[list[value]]`) — Row-major matrix of computed cell values (`include_values`).
- `formulas` (`list[list[str | null]]`) — Row-major matrix of formulas, with `null` for non-formula cells (`include_formulas`).
- `number_formats` (`list[list[str | null]]`) — Row-major matrix of Excel number formats (`include_number_formats`).
- `style_table` (`list[object]`) — Unique shallow style objects used within the range (`include_styles`).
- `style_ids` (`list[list[int]]`) — Row-major matrix of integer indices into `style_table` (`include_styles`).
- `row_hidden` (`list[bool]`) — Boolean flags aligned to the requested rows (`include_hidden_flags`).
- `column_hidden` (`list[bool]`) — Boolean flags aligned to the requested columns (`include_hidden_flags`).
- `merged_ranges` (`list[str]`) — Deduplicated A1 ranges for merges intersecting the requested range (`include_merged_info`).
- `geometry` (`object`) — Range-level layout: `left`, `top`, `width`, `height` (floats, points) (`include_geometry`).

Each `style_table` entry includes:

- `font_name` (`str` | `null`)
- `font_size` (`float` | `null`)
- `font_bold` (`bool` | `null`)
- `font_italic` (`bool` | `null`)
- `font_color` (normalized color value | `null`)
- `horizontal_alignment` (`str` | `null`)
- `vertical_alignment` (`str` | `null`)
- `wrap_text` (`bool` | `null`)
- `fill_color` (normalized color value | `null`)

## Notes

- `formulas[row][col]`, `number_formats[row][col]`, and `style_ids[row][col]` all align to `values[row][col]`.
- `style_table` is shallow and intended for inspection, not full Excel style round-tripping.
- `row_hidden` and `column_hidden` are aligned to the requested range offsets, not absolute worksheet coordinates.
- `merged_ranges` avoids repeating merge metadata for every member cell.

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
