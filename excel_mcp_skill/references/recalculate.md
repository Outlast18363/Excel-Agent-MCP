# recalculate

## Purpose

Triggers a full Excel application recalculation and optionally scans formula cells for error values, returning counts and sample cell locations per error type.

## When to use

Call this when you need an explicit Excel recalculation pass or a quick error audit (`#DIV/0!`, `#REF!`, etc.) without opening each cell manually. `set_range` writes through the live Excel session, and dependent cells may already refresh if Excel is running in automatic calculation mode, so `recalculate` is best treated as the deterministic refresh-and-audit step rather than something that is always required after every edit.

## Parameters

- **workbook_id** (`str`, required) — ID of the open workbook in the live Excel session.
- **scope** (`str`, default `"workbook"`) — Which area is scanned after recalc: `"workbook"`, `"sheet"`, or `"range"`.
- **sheet** (`str | None`, default `None`) — Sheet name; required when `scope` is `"sheet"` or `"range"`.
- **range** (`str | None`, default `None`) — A1-style range; required when `scope` is `"range"` (use with `sheet`).
- **scan_errors** (`bool`, default `True`) — If true, scan formula cells for errors and populate `total_errors` and `error_summary`.
- **return_formula_stats** (`bool`, default `True`) — If true, include `total_formulas` in the response.
- **max_error_locations_per_type** (`int`, default `50`) — Maximum sample locations listed per error type in `error_summary`.

## Response `data` fields

Always present:

- **scope** (`str`) — Echo of the effective scan scope.
- **recalculated** (`bool`) — Always `True` when the tool succeeds.

Conditional (each appears only when the corresponding inputs or flags apply):

- **sheet** (`str`) — Only when the `sheet` parameter is non-null.
- **range** (`str`) — Only when the `range` parameter is non-null.
- **total_formulas** (`int`) — Only when `return_formula_stats=True`.
- **total_errors** (`int`) — Only when `scan_errors=True`.
- **error_summary** (`dict`) — Only when `scan_errors=True`.

## Notes

Recalculation always runs at the Excel application level (`app.calculate()`), regardless of `scope`. The `scope` parameter only limits which cells are scanned afterward: `"workbook"` uses each sheet’s used range, `"sheet"` uses one sheet’s used range (requires `sheet`), and `"range"` limits to the given `sheet` and `range`. Error scanning considers only cells that contain formulas (values whose formula text starts with `=`). For each error type, `locations` is capped at `max_error_locations_per_type`.

`set_range` itself does not call `calculate()`. It writes values/formulas/styles and optionally saves, so whether formulas appear refreshed immediately depends on Excel's own calculation mode. Use `recalculate` when you want the MCP to force a calculation pass and report scan results explicitly.

`error_summary` maps each error string to a count and a list of cell addresses (sheet context implied by scan scope):

```json
{
  "#DIV/0!": {"count": 1, "locations": ["C1"]},
  "#REF!": {"count": 3, "locations": ["D5", "D6", "D7"]}
}
```

## Example

```python
server.recalculate(workbook_id, scope="sheet", sheet="Data")
```
