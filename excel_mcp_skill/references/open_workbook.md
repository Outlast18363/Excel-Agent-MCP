# open_workbook

## Purpose

Opens a live Excel workbook and registers a workbook session. Returns a `workbook_id` and metadata needed for later tools.

## When to use

Call this before any other workbook-scoped tool. Subsequent calls pass the returned `workbook_id`.

## Parameters

- `path` (`str`, required) — Filesystem path to the workbook; resolved to an absolute path internally.
- `read_only` (`bool`, default `False`) — Open without allowing edits.
- `visible` (`bool`, default `False`) — Whether the managed Excel window is shown.
- `create_if_missing` (`bool`, default `False`) — If the file is missing, create a new workbook, save it at `path`, and create parent directories as needed.

## Response `data` fields

- `workbook_id` (`str`) — Handle for all later tool calls.
- `path` (`str`) — Resolved absolute file path.
- `sheet_names` (`list[str]`) — Sheet names in workbook order.
- `active_sheet` (`str`) — Currently active sheet name.
- `read_only` (`bool`) — Whether the session is read-only.

## Notes

- Same resolved path as an existing session reuses that session; no duplicate workbook. The returned session is unchanged even if this call’s `visible` or `read_only` differ from the original open.
- The service keeps a small pool of Excel Application instances keyed by `visible` (one visible, one hidden).

## Example

```python
open_workbook(path="/Users/me/financials.xlsx", visible=False)
```
