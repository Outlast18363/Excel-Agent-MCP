# `get_sheet_state`

## What It Does

Returns sheet-level metadata for one worksheet.

It can include used range boundaries, hidden rows and columns, merged ranges, formula and non-empty counts, plus chart and shape counts.

## When To Use It

Use this before `get_range` when you need fast structural context for an unfamiliar or layout-sensitive sheet.

## Parameters

- `workbook_id: str`
- `sheet: str`
- `include_used_range: bool = True`
- `include_hidden: bool = True`
- `include_merged_ranges: bool = True`
- `include_formula_stats: bool = True`
- `include_object_counts: bool = True`

## Returns In `data`

- `sheet`
- `visible`
- `used_range`
- `max_row`
- `max_col`
- `hidden_rows`
- `hidden_columns`
- `merged_ranges`
- `formula_count`
- `nonempty_cell_count`
- `chart_count`
- `shape_count`

## Notes

- `used_range` means Excel's current occupied rectangle for the sheet, not an edit-history region.
- `hidden_rows` is returned as row numbers.
- `hidden_columns` is returned as column letters such as `["B", "D"]`.
- Formula and non-empty counts are computed from the current used range.

## Example

```python
get_sheet_state(
    workbook_id="wb_001",
    sheet="Income Statement",
)
```
