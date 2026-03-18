"""Excel-facing service layer for the Excel MCP server."""

from __future__ import annotations

import base64
import gzip
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib import error as urllib_error
from urllib import request as urllib_request

from .helpers import (
    ExcelServiceError,
    apply_number_format,
    apply_style,
    build_cell_payload,
    column_number_to_name,
    extract_excel_error,
    get_address,
    get_formula_and_nonempty_counts,
    get_hidden_columns,
    get_hidden_rows,
    get_merged_ranges,
    get_range_geometry,
    normalize_formula_grid,
    normalize_matrix_input,
    normalize_taco_pattern,
    normalize_taco_ref_key,
    safe_count,
    sheet_visible,
    taco_ref_to_a1_address,
    temporary_screenshot_path,
)
from .types import JsonValue

TACO_API_URL = "http://127.0.0.1:4567/api/taco/patterns"


@dataclass(slots=True)
class WorkbookSession:
    """Track a live workbook registered with the MCP process."""

    workbook_id: str
    workbook: Any
    app: Any
    path: str
    read_only: bool
    visible: bool
    trace_graph_sheet: str | None = None
    trace_graph_ready: bool = False
    trace_graph_dirty: bool = True


class ExcelService:
    """Own workbook registry and live Excel interactions for MCP tools."""

    def __init__(self) -> None:
        """Initialize the service registry and app cache."""
        self._workbooks: dict[str, WorkbookSession] = {}
        self._path_index: dict[str, str] = {}
        self._apps: dict[bool, Any] = {}
        self._next_workbook_number = 1
        self._active_trace_graph: tuple[str, str] | None = None

    def open_workbook(
        self,
        *,
        path: str,
        read_only: bool = False,
        visible: bool = True,
        create_if_missing: bool = False,
    ) -> dict[str, JsonValue]:
        xw = self._require_xlwings()
        workbook_path = Path(path).expanduser()
        resolved_path = str(workbook_path.resolve(strict=False))

        existing_id = self._path_index.get(resolved_path)
        if existing_id:
            existing_session = self._workbooks.get(existing_id)
            if existing_session is not None:
                return self._session_payload(existing_session)

        app = self._get_or_create_app(visible=visible)

        if workbook_path.exists():
            workbook = self._find_open_workbook(app=app, target_path=resolved_path)
            if workbook is None:
                workbook = app.books.open(resolved_path, read_only=read_only)
        else:
            if not create_if_missing:
                raise ExcelServiceError(f"Workbook does not exist: {resolved_path}")

            workbook_path.parent.mkdir(parents=True, exist_ok=True)
            workbook = app.books.add()
            workbook.save(resolved_path)

        workbook_id = self._new_workbook_id()
        session = WorkbookSession(
            workbook_id=workbook_id,
            workbook=workbook,
            app=app,
            path=resolved_path,
            read_only=bool(read_only),
            visible=bool(visible),
        )
        self._workbooks[workbook_id] = session
        self._path_index[resolved_path] = workbook_id
        return self._session_payload(session)

    def get_sheet_state(
        self,
        *,
        workbook_id: str,
        sheet: str,
        include_used_range: bool = True,
        include_hidden: bool = True,
        include_merged_ranges: bool = True,
        include_formula_stats: bool = True,
        include_object_counts: bool = True,
    ) -> dict[str, JsonValue]:
        worksheet = self._get_sheet(workbook_id=workbook_id, sheet_name=sheet)
        used_range = worksheet.used_range
        start_row = int(used_range.row)
        start_col = int(used_range.column)
        row_count = int(used_range.rows.count)
        col_count = int(used_range.columns.count)
        max_row = start_row + max(row_count - 1, 0)
        max_col = start_col + max(col_count - 1, 0)

        data: dict[str, JsonValue] = {
            "sheet": worksheet.name,
            "visible": sheet_visible(worksheet),
        }

        if include_used_range:
            data["used_range"] = get_address(used_range)
            data["max_row"] = max_row
            data["max_col"] = max_col

        if include_hidden:
            data["hidden_rows"] = get_hidden_rows(worksheet, start_row, max_row)
            data["hidden_columns"] = get_hidden_columns(worksheet, start_col, max_col)

        if include_merged_ranges:
            data["merged_ranges"] = get_merged_ranges(used_range)

        if include_formula_stats:
            formula_count, nonempty_count = get_formula_and_nonempty_counts(used_range)
            data["formula_count"] = formula_count
            data["nonempty_cell_count"] = nonempty_count

        if include_object_counts:
            data["chart_count"] = safe_count(lambda: int(worksheet.api.ChartObjects().Count))
            data["shape_count"] = safe_count(lambda: int(worksheet.api.Shapes.Count))

        return data

    def get_range(
        self,
        *,
        workbook_id: str,
        sheet: str,
        range_address: str,
        include_values: bool = True,
        include_formulas: bool = False,
        include_number_formats: bool = False,
        include_styles: bool = False,
        include_geometry: bool = False,
        include_hidden_flags: bool = False,
        include_merged_info: bool = False,
    ) -> dict[str, JsonValue]:
        target_range = self._get_range(workbook_id=workbook_id, sheet_name=sheet, range_address=range_address)
        matrix: list[list[dict[str, JsonValue]]] = []
        cells: list[dict[str, JsonValue]] = []

        for row_offset in range(int(target_range.rows.count)):
            matrix_row: list[dict[str, JsonValue]] = []
            for col_offset in range(int(target_range.columns.count)):
                cell = target_range[row_offset, col_offset]
                cell_payload = build_cell_payload(
                    cell=cell,
                    include_values=include_values,
                    include_formulas=include_formulas,
                    include_number_formats=include_number_formats,
                    include_styles=include_styles,
                    include_geometry=include_geometry,
                    include_hidden_flags=include_hidden_flags,
                    include_merged_info=include_merged_info,
                )
                matrix_row.append(cell_payload)
                cells.append(cell_payload)
            matrix.append(matrix_row)

        data: dict[str, JsonValue] = {
            "sheet": sheet,
            "range": get_address(target_range),
            "matrix": matrix,
            "cells": cells,
        }

        if include_geometry:
            data["geometry"] = get_range_geometry(target_range)

        return data

    def set_range(
        self,
        *,
        workbook_id: str,
        sheet: str,
        range_address: str,
        values: Any = None,
        formulas: Any = None,
        number_format: Any = None,
        style: dict[str, Any] | None = None,
        clear_contents: bool = False,
        clear_formats: bool = False,
        save_after: bool = False,
    ) -> dict[str, JsonValue]:
        session = self._get_workbook_session(workbook_id)
        target_range = self._get_range(workbook_id=workbook_id, sheet_name=sheet, range_address=range_address)
        rows = int(target_range.rows.count)
        cols = int(target_range.columns.count)

        values_matrix = (
            normalize_matrix_input(values, rows, cols, "values")
            if values is not None
            else None
        )
        formulas_matrix = (
            normalize_matrix_input(formulas, rows, cols, "formulas")
            if formulas is not None
            else None
        )

        if clear_contents:
            target_range.clear_contents()

        if clear_formats:
            target_range.clear_formats()

        updated_values = False
        if values_matrix is not None:
            target_range.value = values_matrix
            updated_values = True

        updated_formulas = False
        if formulas_matrix is not None:
            target_range.formula = formulas_matrix
            updated_formulas = True

        updated_style = False
        if number_format is not None:
            apply_number_format(target_range, number_format, rows, cols)
            updated_style = True

        if style:
            apply_style(target_range, style)
            updated_style = True

        saved = False
        if save_after:
            session.workbook.save()
            saved = True

        self._mark_trace_graph_dirty(workbook_id)

        return {
            "sheet": sheet,
            "range": get_address(target_range),
            "updated_values": updated_values,
            "updated_formulas": updated_formulas,
            "updated_style": updated_style,
            "saved": saved,
        }

    def recalculate(
        self,
        *,
        workbook_id: str,
        scope: str = "workbook",
        sheet: str | None = None,
        range_address: str | None = None,
        scan_errors: bool = True,
        return_formula_stats: bool = True,
        max_error_locations_per_type: int = 50,
    ) -> dict[str, JsonValue]:
        session = self._get_workbook_session(workbook_id)
        session.app.calculate()
        targets = self._recalc_targets(workbook_id=workbook_id, scope=scope, sheet=sheet, range_address=range_address)

        total_formulas = 0
        error_summary: dict[str, dict[str, JsonValue]] = {}
        total_errors = 0

        if scan_errors or return_formula_stats:
            for target in targets:
                for cell in target:
                    formula = cell.formula
                    if not isinstance(formula, str) or not formula.startswith("="):
                        continue

                    total_formulas += 1
                    if not scan_errors:
                        continue

                    error_literal = extract_excel_error(cell.value)
                    if error_literal is None:
                        continue

                    total_errors += 1
                    bucket = error_summary.setdefault(
                        error_literal,
                        {"count": 0, "locations": []},
                    )
                    bucket["count"] = int(bucket["count"]) + 1
                    locations = bucket["locations"]
                    if (
                        isinstance(locations, list)
                        and len(locations) < max_error_locations_per_type
                    ):
                        locations.append(get_address(cell))

        data: dict[str, JsonValue] = {
            "scope": scope,
            "recalculated": True,
        }

        if sheet is not None:
            data["sheet"] = sheet
        if range_address is not None:
            data["range"] = range_address
        if return_formula_stats:
            data["total_formulas"] = total_formulas
        if scan_errors:
            data["total_errors"] = total_errors
            data["error_summary"] = error_summary

        return data

    def local_screenshot(
        self,
        *,
        workbook_id: str,
        sheet: str,
        range_address: str,
        output_path: str | None = None,
        return_base64: bool = False,
    ) -> dict[str, JsonValue]:
        target_range = self._get_range(workbook_id=workbook_id, sheet_name=sheet, range_address=range_address)
        
        target_path = Path(output_path).expanduser() if output_path else temporary_screenshot_path()
        target_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Use xlwings native to_png method
        target_range.to_png(str(target_path))

        data: dict[str, JsonValue] = {
            "sheet": sheet,
            "range": get_address(target_range),
            "image_path": str(target_path),
        }

        if return_base64:
            data["base64"] = base64.b64encode(target_path.read_bytes()).decode("ascii")

        return data

    def trace_formula(
        self,
        *,
        workbook_id: str,
        sheet: str,
        range_address: str,
        direction: str,
        direct_only: bool = True,
        refresh_graph: bool = True,
    ) -> dict[str, JsonValue]:
        """Trace formula precedents or dependents for a cell or range.

        Parameters:
            workbook_id: The workbook handle returned by ``open_workbook``.
            sheet: The sheet name containing the target range.
            range_address: The A1 target cell or range to trace.
            direction: Either ``precedents`` or ``dependents``.
            direct_only: Whether to return only direct edges.
            refresh_graph: Whether to force a rebuild of the external TACO graph.

        Returns:
            A JSON-safe trace payload containing the queried subgraph.
        """

        session = self._get_workbook_session(workbook_id)
        worksheet = self._get_sheet(workbook_id=workbook_id, sheet_name=sheet)
        target_range = self._get_range(
            workbook_id=workbook_id,
            sheet_name=sheet,
            range_address=range_address,
        )
        normalized_direction = self._normalize_trace_direction(direction)
        build_limit_address, build_limit_row, build_limit_col = self._get_trace_build_limit(worksheet)
        self._validate_trace_target(
            target_range=target_range,
            build_limit_row=build_limit_row,
            build_limit_col=build_limit_col,
        )

        graph_source = "cache"
        if refresh_graph or not self._can_reuse_trace_graph(session, workbook_id, sheet):
            formula_matrix = self._build_trace_formula_matrix(
                workbook_id=workbook_id,
                sheet=sheet,
                build_limit_address=build_limit_address,
            )
            self._post_taco_request(
                {
                    "type": "build",
                    "graph": "taco",
                    "formulae": formula_matrix,
                }
            )
            session.trace_graph_sheet = sheet
            session.trace_graph_ready = True
            session.trace_graph_dirty = False
            self._active_trace_graph = (workbook_id, sheet)
            graph_source = "rebuilt"

        query_response = self._post_taco_request(
            {
                "type": "dep" if normalized_direction == "dependents" else "prec",
                "range": f"{sheet}!{get_address(target_range)}",
                "isDirect": bool(direct_only),
            }
        )

        return {
            "sheet": sheet,
            "range": get_address(target_range),
            "direction": normalized_direction,
            "direct_only": bool(direct_only),
            "graph_source": graph_source,
            "graph_complete": True,
            "subgraph": self._serialize_taco_subgraph(query_response),
        }

    def _require_xlwings(self) -> Any:
        try:
            import xlwings as xw
        except ImportError as exc:
            raise ExcelServiceError(
                "xlwings is required to run the Excel MCP server."
            ) from exc
        return xw

    def _get_or_create_app(self, *, visible: bool) -> Any:
        app = self._apps.get(visible)
        if app is not None:
            return app

        xw = self._require_xlwings()
        app = xw.App(visible=visible, add_book=False)
        self._apps[visible] = app
        return app

    def _new_workbook_id(self) -> str:
        workbook_id = f"wb_{self._next_workbook_number:03d}"
        self._next_workbook_number += 1
        return workbook_id

    def _session_payload(self, session: WorkbookSession) -> dict[str, JsonValue]:
        try:
            active_sheet = session.workbook.app.selection.sheet.name
        except Exception:
            active_sheet = session.workbook.sheets[0].name
            
        return {
            "workbook_id": session.workbook_id,
            "path": session.path,
            "sheet_names": [sheet.name for sheet in session.workbook.sheets],
            "active_sheet": active_sheet,
            "read_only": session.read_only,
        }

    def _find_open_workbook(self, *, app: Any, target_path: str) -> Any | None:
        for workbook in app.books:
            try:
                workbook_path = str(Path(workbook.fullname).resolve(strict=False))
            except Exception:
                continue
            if workbook_path == target_path:
                return workbook
        return None

    def _get_workbook_session(self, workbook_id: str) -> WorkbookSession:
        session = self._workbooks.get(workbook_id)
        if session is None:
            raise ExcelServiceError(f"Unknown workbook id: {workbook_id}")
        return session

    def _get_sheet(self, *, workbook_id: str, sheet_name: str) -> Any:
        workbook = self._get_workbook_session(workbook_id).workbook
        try:
            return workbook.sheets[sheet_name]
        except Exception as exc:
            raise ExcelServiceError(
                f"Sheet `{sheet_name}` was not found in workbook `{workbook_id}`."
            ) from exc

    def _get_range(self, *, workbook_id: str, sheet_name: str, range_address: str) -> Any:
        worksheet = self._get_sheet(workbook_id=workbook_id, sheet_name=sheet_name)
        try:
            return worksheet.range(range_address)
        except Exception as exc:
            raise ExcelServiceError(
                f"Range `{range_address}` is invalid on sheet `{sheet_name}`."
            ) from exc

    def close_all(self) -> None:
        """Safely close all registered workbooks and quit any managed Excel apps."""
        for workbook_id, session in list(self._workbooks.items()):
            try:
                session.workbook.close()
            except Exception:
                pass
        self._workbooks.clear()
        self._path_index.clear()
        
        for visible, app in list(self._apps.items()):
            try:
                app.quit()
            except Exception:
                pass
        self._apps.clear()
        self._active_trace_graph = None

    def _recalc_targets(
        self,
        *,
        workbook_id: str,
        scope: str,
        sheet: str | None,
        range_address: str | None,
    ) -> list[Any]:
        session = self._get_workbook_session(workbook_id)
        if scope == "workbook":
            return [worksheet.used_range for worksheet in session.workbook.sheets]
        if scope == "sheet":
            if not sheet:
                raise ExcelServiceError("`sheet` is required when scope is `sheet`.")
            return [self._get_sheet(workbook_id=workbook_id, sheet_name=sheet).used_range]
        if scope == "range":
            if not sheet:
                raise ExcelServiceError("`sheet` is required when scope is `range`.")
            if not range_address:
                raise ExcelServiceError("`range` is required when scope is `range`.")
            return [self._get_range(workbook_id=workbook_id, sheet_name=sheet, range_address=range_address)]
        raise ExcelServiceError("`scope` must be one of `workbook`, `sheet`, or `range`.")

    def _normalize_trace_direction(self, direction: str) -> str:
        """Validate and normalize the requested trace direction.

        Parameters:
            direction: The user-provided trace direction string.

        Returns:
            The normalized lowercase direction name.
        """

        normalized_direction = direction.strip().lower()
        if normalized_direction not in {"precedents", "dependents"}:
            raise ExcelServiceError("`direction` must be `precedents` or `dependents`.")
        return normalized_direction

    def _can_reuse_trace_graph(
        self,
        session: WorkbookSession,
        workbook_id: str,
        sheet: str,
    ) -> bool:
        """Return whether the currently loaded external graph can be reused.

        Parameters:
            session: The active workbook session.
            workbook_id: The workbook identifier being queried.
            sheet: The sheet name being queried.

        Returns:
            ``True`` when the loaded graph matches the same workbook and sheet and
            the session has not marked it dirty.
        """

        return (
            session.trace_graph_ready
            and not session.trace_graph_dirty
            and session.trace_graph_sheet == sheet
            and self._active_trace_graph == (workbook_id, sheet)
        )

    def _get_trace_build_limit(self, worksheet: Any) -> tuple[str, int, int]:
        """Compute the bottom-right A1 address used to build the sheet graph.

        Parameters:
            worksheet: The live xlwings sheet object to inspect.

        Returns:
            A tuple of ``(address, max_row, max_col)`` for the build rectangle.
        """

        used_range = worksheet.used_range
        start_row = int(used_range.row)
        start_col = int(used_range.column)
        row_count = int(used_range.rows.count)
        col_count = int(used_range.columns.count)
        max_row = max(start_row + row_count - 1, 1)
        max_col = max(start_col + col_count - 1, 1)
        return f"A1:{column_number_to_name(max_col)}{max_row}", max_row, max_col

    def _validate_trace_target(
        self,
        *,
        target_range: Any,
        build_limit_row: int,
        build_limit_col: int,
    ) -> None:
        """Validate that the trace target fits inside the graph build rectangle.

        Parameters:
            target_range: The xlwings range requested by the caller.
            build_limit_row: The inclusive bottom-most row in the graph build.
            build_limit_col: The inclusive right-most column in the graph build.

        Returns:
            ``None``. The function raises when the target falls outside the build.
        """

        target_row = int(target_range.row)
        target_col = int(target_range.column)
        last_row = target_row + int(target_range.rows.count) - 1
        last_col = target_col + int(target_range.columns.count) - 1
        if last_row > build_limit_row or last_col > build_limit_col:
            limit_address = f"A1:{column_number_to_name(build_limit_col)}{build_limit_row}"
            raise ExcelServiceError(
                f"Range `{get_address(target_range)}` lies outside the trace graph build area `{limit_address}`."
            )

    def _build_trace_formula_matrix(
        self,
        *,
        workbook_id: str,
        sheet: str,
        build_limit_address: str,
    ) -> list[list[str]]:
        """Read and normalize the formula grid used to build the TACO graph.

        Parameters:
            workbook_id: The workbook identifier containing the target sheet.
            sheet: The sheet name to read.
            build_limit_address: The A1 rectangle used for graph construction.

        Returns:
            A dense 2D string matrix aligned to worksheet coordinates.
        """

        build_range = self._get_range(
            workbook_id=workbook_id,
            sheet_name=sheet,
            range_address=build_limit_address,
        )
        rows = int(build_range.rows.count)
        cols = int(build_range.columns.count)
        return normalize_formula_grid(build_range.formula, rows, cols)

    def _post_taco_request(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Send a JSON request to the local TACO-Lens backend.

        Parameters:
            payload: The JSON body to post to the backend.

        Returns:
            The decoded JSON response from the backend.
        """

        request_body = json.dumps(payload).encode("utf-8")
        request = urllib_request.Request(
            TACO_API_URL,
            data=request_body,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib_request.urlopen(request, timeout=30) as response:
                body = response.read()
                if "gzip" in response.headers.get("Content-Encoding", "").lower():
                    try:
                        body = gzip.decompress(body)
                    except OSError:
                        pass
        except urllib_error.HTTPError as exc:
            raise ExcelServiceError(
                f"TACO backend request failed with HTTP {exc.code}."
            ) from exc
        except urllib_error.URLError as exc:
            raise ExcelServiceError(
                "Trace formula requires the local TACO-Lens backend at "
                f"`{TACO_API_URL}` to be running."
            ) from exc

        try:
            decoded = body.decode("utf-8")
            data = json.loads(decoded)
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ExcelServiceError("TACO backend returned an unreadable response.") from exc

        if not isinstance(data, dict):
            raise ExcelServiceError("TACO backend returned an unexpected response shape.")
        return data

    def _serialize_taco_subgraph(
        self,
        response: dict[str, Any],
    ) -> dict[str, list[dict[str, JsonValue]]]:
        """Convert a raw TACO subgraph response into stable MCP-facing output.

        Parameters:
            response: The decoded JSON response returned by the TACO backend.

        Returns:
            A normalized mapping from source ranges to traced edge payloads.
        """

        taco_payload = response.get("taco", {})
        if not isinstance(taco_payload, dict):
            return {}
        raw_subgraph = taco_payload.get("default-sheet-name", {})
        if not isinstance(raw_subgraph, dict):
            return {}

        normalized_subgraph: dict[str, list[dict[str, JsonValue]]] = {}
        for raw_key, raw_edges in raw_subgraph.items():
            if not isinstance(raw_edges, list):
                continue
            normalized_edges: list[dict[str, JsonValue]] = []
            for raw_edge in raw_edges:
                if not isinstance(raw_edge, dict):
                    continue
                ref_payload = raw_edge.get("ref")
                edge_meta = raw_edge.get("edgeMeta")
                if not isinstance(ref_payload, dict) or not isinstance(edge_meta, dict):
                    continue
                normalized_edges.append(
                    {
                        "range": taco_ref_to_a1_address(ref_payload),
                        "pattern": normalize_taco_pattern(edge_meta.get("patternType")),
                    }
                )
            if normalized_edges:
                normalized_subgraph[normalize_taco_ref_key(str(raw_key))] = sorted(
                    normalized_edges,
                    key=lambda edge: (str(edge["range"]), str(edge["pattern"])),
                )

        return dict(sorted(normalized_subgraph.items()))

    def _mark_trace_graph_dirty(self, workbook_id: str) -> None:
        """Mark the trace graph state for a workbook session as stale.

        Parameters:
            workbook_id: The workbook identifier whose cached trace state changed.

        Returns:
            ``None``. The in-memory freshness metadata is updated in place.
        """

        session = self._workbooks.get(workbook_id)
        if session is None:
            return
        session.trace_graph_dirty = True
        if self._active_trace_graph and self._active_trace_graph[0] == workbook_id:
            self._active_trace_graph = None


excel_service = ExcelService()
