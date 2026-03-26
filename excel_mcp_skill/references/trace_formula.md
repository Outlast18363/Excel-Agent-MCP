# `trace_formula`

## What It Does

Traces formula precedents or dependents for a cell or range using the in-process `formulas` workbook model.

It can return a direct graph slice, a bounded-depth graph slice, or a fully transitive graph depending on `max_depth`.

## When To Use It

Use this when you need dependency reasoning, impact analysis, or formula lineage instead of just cell contents.

## Parameters

- `workbook_id: str`
- `sheet: str`
- `range: str`
- `direction: str`
  - `precedents` for upstream inputs
  - `dependents` for downstream formulas
- `max_depth: int | None = 1`
- `include_addresses: bool = True`

## Returns In `data`

- `sheet`
- `range`
- `direction`
- `max_depth`
- `complete`
- `nodes`
- `edges`

Each graph edge looks like:

```json
{
  "from": "Inputs!A1:A2",
  "to": "B2"
}
```

## Notes

- This tool does not require Docker, Java, or any local HTTP backend.
- The implementation snapshots the live workbook and builds a native workbook graph with `formulas.ExcelModel`.
- Cross-sheet refs are preserved in normalized ids like `Inputs!A1` when they leave the active sheet.
- `max_depth=1` means direct edges only.
- `max_depth=None` means full transitive traversal.

## Example

```python
trace_formula(
    workbook_id="wb_001",
    sheet="Calc",
    range="B2",
    direction="dependents",
    max_depth=1,
    include_addresses=True,
)
```
