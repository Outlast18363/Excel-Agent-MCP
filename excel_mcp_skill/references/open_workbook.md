# `open_workbook`

## What It Does

Opens a workbook into the MCP-managed Excel session and returns a `workbook_id`.

If the same resolved path is already registered in the current MCP process, the existing session is reused instead of opening a duplicate workbook.

## When To Use It

Use this first. All other workbook tools require the returned `workbook_id`.

## Parameters

- `path: str`
  - Full filesystem path to the workbook.
- `read_only: bool = False`
  - Open without allowing edits.
- `visible: bool = True`
  - Whether the managed Excel app should be visible on screen.
- `create_if_missing: bool = False`
  - Create and save a new workbook at `path` if the file does not exist.

## Returns In `data`

- `workbook_id`
- `path`
- `sheet_names`
- `active_sheet`
- `read_only`

## Notes

- Session reuse is path-based, so a later call with different `visible` or `read_only` values may still return the existing session.
- `create_if_missing=True` creates a real workbook file on disk at the given path.

## Example

```python
open_workbook(
    path="/Users/me/financials.xlsx",
    read_only=False,
    visible=False,
    create_if_missing=False,
)
```
