# get_sheet_state

## Purpose

Returns objective sheet-level metadata for a single worksheet in an open workbook: used range bounds, visibility of rows/columns, merged areas, formula and nonempty cell counts, and embedded object counts.

## When to use

Use this when you need a factual snapshot of sheet structure and content density before reading cells, tracing formulas, or editing ranges. Prefer it over scanning the grid when you only need bounds, hidden geometry, or aggregate counts.

## Parameters

- **workbook_id** (`str`, required): Open workbook identifier from `open_workbook`.
- **sheet** (`str`, required): Worksheet name.
- **include_used_range** (`bool`, default `True`): Include `used_range`, `max_row`, and `max_col`.
- **include_hidden** (`bool`, default `True`): Include `hidden_rows` and `hidden_columns` (scoped to the used range).
- **include_merged_ranges** (`bool`, default `True`): Include `merged_ranges`.
- **include_formula_stats** (`bool`, default `True`): Include `formula_count` and `nonempty_cell_count`.
- **include_object_counts** (`bool`, default `True`): Include `chart_count` and `shape_count`.

## Response `data` fields

Always present:

- **sheet** (`str`): The worksheet name.
- **visible** (`bool`): Whether the sheet is visible in Excel (sheet tab visibility), not whether the application window is visible.

If `include_used_range` is true:

- **used_range** (`str`): A1-style address of Excel’s occupied rectangle for the sheet.
- **max_row** (`int`): Last occupied row (1-based).
- **max_col** (`int`): Last occupied column (1-based numeric index).

If `include_hidden` is true:

- **hidden_rows** (`list[int]`): 1-based row indices of hidden rows within the used range.
- **hidden_columns** (`list[str]`): Column letters of hidden columns within the used range (e.g. `["B", "D"]`).

If `include_merged_ranges` is true:

- **merged_ranges** (`list[str]`): Deduplicated A1 addresses of merged cell areas.

If `include_formula_stats` is true:

- **formula_count** (`int`): Count of cells whose stored formula starts with `=`.
- **nonempty_cell_count** (`int`): Count of cells with a non-null, non-empty value.

If `include_object_counts` is true:

- **chart_count** (`int`): Number of charts on the sheet.
- **shape_count** (`int`): Number of shapes on the sheet.

## Notes

Each optional block above is omitted when its matching `include_*` argument is false; only `sheet` and `visible` are always returned. Hidden row and column lists are limited to the sheet’s used range, not the entire grid. `merged_ranges` is deduplicated. Formula statistics count by cell semantics as implemented in the service (formula prefix `=`, nonempty value definition as in the engine).

## Example

```python
response = server.get_sheet_state(workbook_id, "Data")
data = response["data"]
# e.g. data["sheet"] == "Data", data["max_row"] == 10, data["max_col"] == 3
# data["formula_count"] == 11, data["nonempty_cell_count"] == 20
```
