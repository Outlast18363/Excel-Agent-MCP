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
    """Open a live Excel workbook and register a workbook session.

    Parameters:
        path: Full file path of the workbook to open.
        read_only: Whether to open the workbook without allowing edits.
        visible: Whether Excel should be visible on screen.
        create_if_missing: Whether to create a new workbook when the file is missing.

    Returns:
        A shared MCP response containing workbook session metadata.
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
    """Return objective sheet-level metadata for a workbook sheet.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name to inspect.
        include_used_range: Whether to include used range boundaries.
        include_hidden: Whether to include hidden rows and columns.
        include_merged_ranges: Whether to include merged-cell ranges.
        include_formula_stats: Whether to include formula and non-empty cell counts.
        include_object_counts: Whether to include chart and shape counts.

    Returns:
        A shared MCP response containing structural sheet metadata.
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
    """Return a matrix and flattened cell view for an explicit Excel range.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name containing the target range.
        range: The A1 range to inspect.
        include_values: Whether to include cell values.
        include_formulas: Whether to include formula strings.
        include_number_formats: Whether to include number format strings.
        include_styles: Whether to include shallow style details.
        include_geometry: Whether to include range and cell geometry.
        include_hidden_flags: Whether to include hidden row or column flags.
        include_merged_info: Whether to include merged-cell membership details.

    Returns:
        A shared MCP response containing range-level cell data.
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
    """Write values, formulas, and formatting into an explicit Excel range.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name containing the target range.
        range: The A1 range to modify.
        values: Optional 2D values to write.
        formulas: Optional 2D formulas to write.
        number_format: Optional scalar or 2D number format payload.
        style: Optional style settings to apply to the target range.
        clear_contents: Whether to clear cell contents before writing.
        clear_formats: Whether to clear formatting before applying updates.
        save_after: Whether to save the workbook after the operation.

    Returns:
        A shared MCP response describing what changed.
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
    """Trigger Excel recalculation and optionally scan formula cells for errors.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        scope: The recalculation scope: ``workbook``, ``sheet``, or ``range``.
        sheet: The target sheet when scope is ``sheet`` or ``range``.
        range: The target range when scope is ``range``.
        scan_errors: Whether to inspect formula cells for Excel errors.
        return_formula_stats: Whether to include total formula counts.
        max_error_locations_per_type: The cap on returned example error locations.

    Returns:
        A shared MCP response containing recalculation and error scan results.
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
    """Capture the rendered on-screen appearance of an Excel range.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name containing the target range.
        range: The A1 range to capture.
        output_path: Optional destination path for the PNG file.
        return_base64: Whether to include base64 PNG bytes in the response.

    Returns:
        A shared MCP response containing screenshot output details.
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
    direct_only: bool = True,
    refresh_graph: bool = True,
) -> McpResponse:
    """Trace formula precedents or dependents for a cell or range.

    Parameters:
        workbook_id: The workbook handle returned by ``open_workbook``.
        sheet: The sheet name containing the target range.
        range: The A1 cell or range to trace.
        direction: Either ``precedents`` or ``dependents``.
        direct_only: Whether to return only direct edges.
        refresh_graph: Whether to rebuild the dependency graph before querying.

    Returns:
        A shared MCP response containing the traced dependency subgraph.
    """

    return _execute_tool(
        lambda: excel_service.trace_formula(
            workbook_id=workbook_id,
            sheet=sheet,
            range_address=range,
            direction=direction,
            direct_only=direct_only,
            refresh_graph=refresh_graph,
        )
    )
