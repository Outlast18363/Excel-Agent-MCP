# trace_formula

## Purpose

Trace formula precedents or dependents for a cell or range by building a native workbook graph from a snapshot of the live file.

## When to use

Use this when you need dependency structure (what feeds a formula, or what depends on a cell) beyond a single cell’s formula text. Prefer it over manual parsing when ranges, cross-sheet refs, or multi-hop chains matter.

## Parameters

- **workbook_id** (str, required) — Open workbook session id.
- **sheet** (str, required) — Sheet name containing the traced range.
- **range** (str, required) — A1 address of the cell or range to trace.
- **direction** (str, required) — `"precedents"` (inputs to formulas) or `"dependents"` (cells that use the target).
- **max_depth** (int | null, default `1`) — `1` = direct edges only; `null` = full transitive traversal.
- **include_addresses** (bool, default `True`) — If true, each node includes `sheet` and `range`; if false, nodes only include `id`.

## Response `data` fields

- **sheet** (str) — Sheet passed in for the trace target.
- **range** (str) — A1 address of the traced target.
- **direction** (str) — `"precedents"` or `"dependents"`.
- **max_depth** (int | null) — Depth limit applied (`null` means unlimited / full traversal).
- **complete** (bool) — Always `True` in the current implementation.
- **nodes** (list) — Graph vertices; shape depends on `include_addresses` (see below).
- **edges** (list) — Directed edges `{ "from": "<id>", "to": "<id>" }`.

**Node shape when `include_addresses=True`:** `{ "id": "B2", "sheet": "Calc", "range": "B2" }`. Same-sheet refs use an `id` without a sheet prefix (e.g. `"B2"`). Cross-sheet refs use normalized ids like `"Inputs!A1"`.

**Node shape when `include_addresses=False`:** `{ "id": "B2" }` only.

**Edge shape:** `{ "from": "Inputs!A1:A2", "to": "B2" }`.

## Notes

- Snapshots the live workbook to a temp file and builds the graph with `formulas.ExcelModel`. No Docker, Java, or local HTTP backend is required.
- Same-sheet references omit the sheet prefix in node `id`; cross-sheet references use `Sheet!Range` in the id where applicable.
- **Precedents:** range references in formulas are preserved on edges (e.g. `=SUM(A1:A2)` yields an edge from `Inputs!A1:A2` to the formula cell).
- **Dependents:** for a single-cell query, ranges that contain that cell are expanded so formulas referencing the range are found as dependents of that cell.
- `max_depth=1` limits to direct edges; `max_depth=None` walks the full transitive closure (subject to the graph).

## Example

```python
server.trace_formula(workbook_id, "Calc", "B2", "precedents", max_depth=1)
```

Direct precedent of `B2` can appear as edge `("Inputs!A1:A2", "B2")` when the formula sums that range.
