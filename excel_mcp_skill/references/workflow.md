# Excel MCP Workflow

This workflow gives Codex a reliable method for solving Excel tasks with the MCP tools in this repository.

The goal is to understand the workbook first, act narrowly, validate after changes, and escalate to visual or dependency tools only when needed.

## Core Principles

- Start from workbook state, not assumptions.
- Read narrowly before writing broadly.
- Prefer the smallest range that gives enough context.
- After edits, validate with recalculation before concluding the task.
- Use screenshots for appearance questions and tracing for dependency questions.
- Keep changes intentional and reversible where possible.

## Default Workflow

### 1. Open The Workbook

Start by calling `open_workbook`.

Use it to establish a live session and obtain the `workbook_id` needed for all later steps.

Choose `read_only=True` when the task is purely investigative. Choose editable mode when the task may require changes.

### 2. Orient Yourself At The Sheet Level

Use `get_sheet_state` on the most relevant sheet before inspecting specific cells.

This is the default way to learn:

- where the real content is
- whether the sheet has hidden structure
- whether merged cells may affect interpretation
- whether the sheet looks formula-heavy or mostly input-driven

Skip or minimize this step only when the task already names a precise range and the sheet is well understood.

### 3. Inspect The Target Region

Use `get_range` to read the exact area you plan to reason about or modify.

Choose options based on the task:

- values for content checks
- formulas for logic checks
- number formats and styles for presentation checks
- geometry or merged info for layout-sensitive work

Read enough surrounding context to avoid isolated edits that break nearby structure.

### 4. Decide The Nature Of The Task

Before editing, classify the problem loosely:

- content task: change inputs, labels, or static values
- formula task: fix or extend spreadsheet logic
- formatting task: improve visual presentation
- diagnostic task: explain how a result is produced or what will be affected by a change

This determines which tools matter most in the next steps.

### 5. Make Narrow, Explicit Edits

Use `set_range` for changes.

Typical uses:

- write values
- write formulas
- clear a block before rewriting it
- apply number formats
- apply lightweight styles

Prefer targeted edits over large blind rewrites. If the requested change spans many cells, still reason about the intended structure before writing.

### 6. Recalculate And Check Health

After changing values or formulas, call `recalculate`.

Use a scope that matches the likely blast radius:

- `range` for tightly localized edits
- `sheet` for sheet-level changes
- `workbook` for cross-sheet or uncertain impact

If errors appear, inspect the affected cells with `get_range` before making further edits.

### 7. Validate Visually When Needed

Use `local_screenshot` when the task depends on appearance rather than only data.

Typical cases:

- reports or dashboards
- spacing, alignment, or header layout
- formatted exports for human review

Do not use screenshots as the default validation method for purely logical tasks.

### 8. Trace Dependencies When Logic Matters

Use `trace_formula` when you need to understand relationships rather than just cell contents.

Typical cases:

- what feeds this output
- what downstream cells depend on this input
- whether a change has broader impact than expected

Use `max_depth=1` first when possible. Use larger depths or `max_depth=None` only when the task needs lineage or impact analysis across multiple steps.

## Common Patterns

### Unknown Workbook Or Sheet

1. `open_workbook`
2. `get_sheet_state`
3. `get_range`

### Fix A Formula Or Broken Model

1. `open_workbook`
2. `get_range` on the suspect area, usually with formulas included
3. `trace_formula` if upstream or downstream logic is unclear
4. `set_range`
5. `recalculate`
6. `get_range` again to confirm the result

### Update Inputs Or Labels

1. `open_workbook`
2. `get_range`
3. `set_range`
4. `recalculate` if formulas depend on the edited cells

### Improve Formatting Or Presentation

1. `open_workbook`
2. `get_range` with styles or geometry as needed
3. `set_range`
4. `local_screenshot`

### Explain A Number Or Impact Of A Change

1. `open_workbook`
2. `get_range`
3. `trace_formula`
4. `get_range` on any important precedent or dependent regions

## Decision Heuristics

- If you do not know where the real content is, start with `get_sheet_state`.
- If you know the location but not the logic, start with `get_range`.
- If the question is about appearance, bring in `local_screenshot`.
- If the question is about causality or downstream risk, bring in `trace_formula`.
- If you edited formulas or inputs, `recalculate` is usually mandatory.
- If a task can be solved by reading a small range, do not jump to workbook-wide actions.

## Completion Standard

A task is usually ready to conclude when:

- the relevant cells were inspected
- any requested edits were applied intentionally
- recalculation was checked when logic changed
- visual output was checked when presentation mattered
- dependency tracing was used when explanation or impact analysis mattered
