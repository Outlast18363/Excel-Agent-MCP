"""MCP server surface for the Excel MCP toolkit."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from mcp.server.fastmcp import FastMCP

from .service import ExcelServiceError, excel_service
from .types import McpResponse, error_response, success_response

mcp_server = FastMCP("Excel MCP", json_response=True)


def _execute_tool(operation: Callable[[], dict[str, Any]]) -> McpResponse:
    """Run a service-layer operation and normalize failures for MCP callers.

    Parameters:
        operation: A zero-argument callable that performs the requested work.

    Returns:
        A shared response envelope containing either data or a user-facing error.
    """

    try:
        return success_response(operation())
    except ExcelServiceError as exc:
        return error_response(str(exc))
    except Exception as exc:  # pragma: no cover - defensive integration guard
        return error_response(f"Unexpected server error: {exc}")


@mcp_server.tool()
def open_workbook(
    path: str,
    read_only: bool = False,
    visible: bool = True,
    create_if_missing: bool = False,
) -> McpResponse:
    """Open a live Excel workbook and return a session ``workbook_id`` for all later calls.

    If the resolved path matches an existing session, the session is reused
    unchanged (even if ``visible`` or ``read_only`` differ).

    Parameters:
        path: Filesystem path to the workbook (resolved to absolute internally).
        read_only: Open without allowing edits.
        visible: Whether the managed Excel window is shown on screen.
        create_if_missing: Create and save a new workbook at ``path`` (with
            parent directories) when the file does not exist.

    Returns:
        ``workbook_id``, resolved ``path``, ``sheet_names``, ``active_sheet``,
        and ``read_only`` flag.
    """

    return _execute_tool(
        lambda: excel_service.open_workbook(
            path=path,
            read_only=read_only,
            visible=visible,
            create_if_missing=create_if_missing,
        )
    )


@mcp_server.tool()
def get_sheet_state(
    workbook_id: str,
    sheet: str,
    include_used_range: bool = True,
    include_hidden: bool = True,
    include_merged_ranges: bool = True,
    include_formula_stats: bool = True,
    include_object_counts: bool = True,
) -> McpResponse:
    """Return sheet-level structural metadata: bounds, hidden rows/columns,
    merged areas, formula/nonempty counts, and chart/shape counts.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name to inspect.
        include_used_range: Return ``used_range`` address, ``max_row``, and
            ``max_col`` for the occupied rectangle.
        include_hidden: Return lists of hidden row numbers and column letters
            within the used range.
        include_merged_ranges: Return deduplicated A1 addresses of merged areas.
        include_formula_stats: Return ``formula_count`` (cells starting with
            ``=``) and ``nonempty_cell_count``.
        include_object_counts: Return ``chart_count`` and ``shape_count``.

    Returns:
        Always includes ``sheet`` and ``visible`` (sheet tab visibility).
        Other fields are conditional on the ``include_*`` flags.
    """

    return _execute_tool(
        lambda: excel_service.get_sheet_state(
            workbook_id=workbook_id,
            sheet=sheet,
            include_used_range=include_used_range,
            include_hidden=include_hidden,
            include_merged_ranges=include_merged_ranges,
            include_formula_stats=include_formula_stats,
            include_object_counts=include_object_counts,
        )
    )


@mcp_server.tool()
def get_range(
    workbook_id: str,
    sheet: str,
    range: str,
    include_values: bool = True,
    include_formulas: bool = False,
    include_number_formats: bool = False,
    include_styles: bool = False,
    include_geometry: bool = False,
    include_hidden_flags: bool = False,
    include_merged_info: bool = False,
) -> McpResponse:
    """Read cell data for an A1 range, returned as a row-major ``matrix`` and
    a flat ``cells`` list (same objects, different layout).

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name containing the target range.
        range: The A1-style address to read (e.g. ``B4:E12``).
        include_values: Return each cell's computed value.
        include_formulas: Return formula strings (``=``-prefixed) where
            present; ``null`` for non-formula cells.
        include_number_formats: Return Excel number format codes.
        include_styles: Return per-cell style snapshots (font, fill, alignment,
            wrap).
        include_geometry: Return per-cell and range-level position/size in
            points (``left``, ``top``, ``width``, ``height``).
        include_hidden_flags: Return ``row_hidden`` and ``column_hidden`` bools.
        include_merged_info: Return ``is_merged`` flag and ``merged_range``
            address when the cell belongs to a merge.

    Returns:
        ``sheet``, ``range``, ``matrix``, ``cells``. Each cell always has
        ``address``, ``row``, ``column``; other fields depend on flags.
    """

    return _execute_tool(
        lambda: excel_service.get_range(
            workbook_id=workbook_id,
            sheet=sheet,
            range_address=range,
            include_values=include_values,
            include_formulas=include_formulas,
            include_number_formats=include_number_formats,
            include_styles=include_styles,
            include_geometry=include_geometry,
            include_hidden_flags=include_hidden_flags,
            include_merged_info=include_merged_info,
        )
    )


@mcp_server.tool()
def set_range(
    workbook_id: str,
    sheet: str,
    range: str,
    values: Any = None,
    formulas: Any = None,
    number_format: Any = None,
    style: dict[str, Any] | None = None,
    clear_contents: bool = False,
    clear_formats: bool = False,
    save_after: bool = False,
) -> McpResponse:
    """Write values, formulas, number formats, and styles into an A1 range.

    Execution order: clear_contents, clear_formats, values, formulas,
    number_format, style, save. If both ``values`` and ``formulas`` are
    given, formulas are written last and win.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name containing the target range.
        range: The A1-style address to modify (e.g. ``D1``, ``A1:B3``).
        values: 2D list matching range shape, or a scalar for a single-cell
            range (auto-wrapped to ``[[value]]``).
        formulas: 2D list matching range shape, or a scalar for a single-cell
            range.
        number_format: A single Excel format string applied uniformly, or a 2D
            list matching range shape for per-cell formats.
        style: Dict of style keys to apply. Supported: ``fill_color``,
            ``font_name``, ``font_size``, ``font_bold``, ``font_italic``,
            ``font_color``, ``horizontal_alignment``, ``vertical_alignment``,
            ``wrap_text``. Colors use ``#RRGGBB`` hex.
        clear_contents: Clear cell contents before writing.
        clear_formats: Clear formatting before applying updates.
        save_after: Save the workbook after the operation.

    Returns:
        ``sheet``, ``range``, and boolean flags ``updated_values``,
        ``updated_formulas``, ``updated_style``, ``saved``.
    """

    return _execute_tool(
        lambda: excel_service.set_range(
            workbook_id=workbook_id,
            sheet=sheet,
            range_address=range,
            values=values,
            formulas=formulas,
            number_format=number_format,
            style=style,
            clear_contents=clear_contents,
            clear_formats=clear_formats,
            save_after=save_after,
        )
    )


@mcp_server.tool()
def recalculate(
    workbook_id: str,
    scope: str = "workbook",
    sheet: str | None = None,
    range: str | None = None,
    scan_errors: bool = True,
    return_formula_stats: bool = True,
    max_error_locations_per_type: int = 50,
) -> McpResponse:
    """Force Excel to recalculate and optionally scan formula cells for errors.

    Recalculation always runs at the application level. The ``scope``
    parameter controls which cells are **scanned** afterward for errors and
    formula counts.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        scope: Scan scope after recalc: ``workbook`` (all sheets),
            ``sheet`` (one sheet), or ``range`` (one range).
        sheet: Required when scope is ``sheet`` or ``range``.
        range: Required when scope is ``range``.
        scan_errors: Inspect formula cells for Excel error values (``#DIV/0!``,
            ``#REF!``, etc.) and return ``total_errors`` and
            ``error_summary``.
        return_formula_stats: Return ``total_formulas`` count.
        max_error_locations_per_type: Cap on sample cell addresses per error
            type in ``error_summary``.

    Returns:
        Always ``scope`` and ``recalculated``. Conditionally ``sheet``,
        ``range``, ``total_formulas``, ``total_errors``, ``error_summary``
        depending on inputs and flags.
    """

    return _execute_tool(
        lambda: excel_service.recalculate(
            workbook_id=workbook_id,
            scope=scope,
            sheet=sheet,
            range_address=range,
            scan_errors=scan_errors,
            return_formula_stats=return_formula_stats,
            max_error_locations_per_type=max_error_locations_per_type,
        )
    )


@mcp_server.tool()
def local_screenshot(
    workbook_id: str,
    sheet: str,
    range: str,
    output_path: str | None = None,
    return_base64: bool = False,
) -> McpResponse:
    """Export a rendered PNG of an Excel range (not a full-screen capture).

    The output is composited onto a white background so every pixel is
    fully opaque.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name containing the target range.
        range: The A1-style address to capture.
        output_path: Destination path for the PNG file. A temporary file is
            created when omitted.
        return_base64: Include base64-encoded PNG bytes in the response as
            ``base64``.

    Returns:
        ``sheet``, ``range``, ``image_path``, and optionally ``base64``.
    """

    return _execute_tool(
        lambda: excel_service.local_screenshot(
            workbook_id=workbook_id,
            sheet=sheet,
            range_address=range,
            output_path=output_path,
            return_base64=return_base64,
        )
    )


@mcp_server.tool()
def trace_formula(
    workbook_id: str,
    sheet: str,
    range: str,
    direction: str,
    max_depth: int | None = 1,
    include_addresses: bool = True,
) -> McpResponse:
    """Trace formula precedents or dependents for a cell or range and return
    a directed graph of nodes and edges.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name containing the target cell or range.
        range: The A1-style address to trace.
        direction: ``precedents`` (upstream inputs) or ``dependents``
            (downstream formulas).
        max_depth: Traversal depth limit. ``1`` = direct edges only;
            ``None`` = full transitive closure.
        include_addresses: When true, each node includes ``sheet`` and
            ``range`` fields alongside ``id``.

    Returns:
        ``sheet``, ``range``, ``direction``, ``max_depth``, ``complete``,
        ``nodes``, and ``edges``. Same-sheet refs omit the sheet prefix in
        node ids; cross-sheet refs use ``Sheet!Range`` format.
    """

    return _execute_tool(
        lambda: excel_service.trace_formula(
            workbook_id=workbook_id,
            sheet=sheet,
            range_address=range,
            direction=direction,
            max_depth=max_depth,
            include_addresses=include_addresses,
        )
    )
