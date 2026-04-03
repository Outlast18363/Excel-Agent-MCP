# set_range

## Purpose

Writes values, formulas, and formatting into an explicit Excel range on an open workbook.

## When to use

Use this tool when you need to change cell contents, formulas, number formats, or visual style for a known A1 range. Prefer it over ad-hoc edits when the target region is fixed and you want predictable batch updates.

## Parameters

- **workbook_id** (`str`, required): Open workbook identifier from `open_workbook`.
- **sheet** (`str`, required): Sheet name.
- **range** (`str`, required): A1-style range (e.g. `D1`, `A1:B3`).
- **values** (`Any`, default `None`): 2D matrix of cell values matching the range shape; for a single-cell range, a scalar is accepted and wrapped to `[[value]]`.
- **formulas** (`Any`, default `None`): 2D matrix of formulas matching the range shape; same scalar rule as `values` for one cell.
- **number_format** (`Any`, default `None`): Either one Excel number-format string applied to the whole range, or a 2D matrix of strings matching the range shape for per-cell formats.
- **style** (`dict[str, Any] | None`, default `None`): Optional style map (see Supported `style` keys).
- **clear_contents** (`bool`, default `False`): If true, clears cell contents before other writes.
- **clear_formats** (`bool`, default `False`): If true, clears formatting before other writes.
- **save_after** (`bool`, default `False`): If true, persists the workbook after updates.

## Supported `style` keys

- **fill_color**: Hex string `#RRGGBB` (fill background).
- **font_name**: Font family name (`str`).
- **font_size**: Point size (`float`).
- **font_bold**: Bold (`bool`).
- **font_italic**: Italic (`bool`).
- **font_color**: Hex string `#RRGGBB`.
- **horizontal_alignment**: Alignment string (e.g. `left`, `center`, `right`).
- **vertical_alignment**: Vertical alignment string (e.g. `top`, `center`, `bottom`).
- **wrap_text**: Wrap text in cell (`bool`).

## Response `data` fields

- **sheet**: Sheet name that was updated.
- **range**: The A1 address that was written.
- **updated_values**: Whether value data was written.
- **updated_formulas**: Whether formula data was written.
- **updated_style**: Whether number format and/or `style` was applied.
- **saved**: Whether the workbook was saved (`save_after`).

## Notes

Execution order: (1) `clear_contents`, (2) `clear_formats`, (3) `values`, (4) `formulas`, (5) `number_format`, (6) `style`, (7) `save_after` if true. Multi-cell `values` and `formulas` must be 2D lists whose dimensions match the target range exactly. If both `values` and `formulas` are provided, formulas are written after values, so formulas win for overlapping cells. `fill_color` and `font_color` use `#RRGGBB` only (no alpha in this map).

## Example

```python
server.set_range(
    workbook_id,
    "Data",
    "D1",
    values=[["Updated"]],
    style={"font_bold": True},
    save_after=True,
)
```
