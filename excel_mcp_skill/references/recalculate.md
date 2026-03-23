# `recalculate`

## What It Does

Triggers Excel recalculation and can scan formula cells for Excel errors afterward.

Supported scopes are:

- `workbook`
- `sheet`
- `range`

## When To Use It

Use this after edits, especially formula edits, before trusting the updated workbook state.

## Parameters

- `workbook_id: str`
- `scope: str = "workbook"`
- `sheet: str | None = None`
- `range: str | None = None`
- `scan_errors: bool = True`
- `return_formula_stats: bool = True`
- `max_error_locations_per_type: int = 50`

## Returns In `data`

- `scope`
- `recalculated`
- `sheet`
- `range`
- `total_formulas`
- `total_errors`
- `error_summary`

## Notes

- The tool always triggers calculation through the live Excel app.
- `scope` decides what gets scanned and summarized in the returned payload.
- Error scanning inspects only formula-bearing cells.
- `error_summary` groups errors by literal such as `#REF!` or `#N/A` and includes sample addresses.

## Example

```python
recalculate(
    workbook_id="wb_001",
    scope="range",
    sheet="Model",
    range="F10:H40",
    scan_errors=True,
)
```
