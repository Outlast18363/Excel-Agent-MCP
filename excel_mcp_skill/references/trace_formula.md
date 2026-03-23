# `trace_formula`

## What It Does

Traces formula precedents or dependents for a cell or range using the local TACO backend.

It can return either direct edges only or a larger transitive dependency subgraph.

## When To Use It

Use this when you need dependency reasoning, impact analysis, or formula lineage instead of just cell contents.

## Parameters

- `workbook_id: str`
- `sheet: str`
- `range: str`
- `direction: str`
  - `precedents` for upstream inputs
  - `dependents` for downstream formulas
- `direct_only: bool = True`
- `refresh_graph: bool = True`

## Returns In `data`

- `sheet`
- `range`
- `direction`
- `direct_only`
- `graph_source`
- `graph_complete`
- `subgraph`

Each subgraph edge looks like:

```json
{
  "range": "B2:B3",
  "pattern": "..."
}
```

## Notes

- This tool requires the local TACO-Lens backend at `http://127.0.0.1:4567/api/taco/patterns`.
- The current implementation builds and queries one worksheet graph at a time.
- The build area is bounded by the sheet's current used-range extent, normalized to an `A1:<bottom_right>` rectangle.
- If the workbook changed after the last graph build, a cached graph may be treated as stale and rebuilt even when `refresh_graph=False`.

## Example

```python
trace_formula(
    workbook_id="wb_001",
    sheet="Calc",
    range="B2",
    direction="dependents",
    direct_only=True,
    refresh_graph=True,
)
```
