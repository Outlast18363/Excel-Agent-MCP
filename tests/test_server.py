"""Unit tests for the MCP tool wrappers exposed by the Excel server."""

from __future__ import annotations

import importlib.util
import os
import unittest
from unittest.mock import patch

from excel_mcp.helpers import ExcelServiceError

MCP_AVAILABLE = importlib.util.find_spec("mcp") is not None

if MCP_AVAILABLE:
    from excel_mcp import server


def unwrap_tool_response(response: object) -> dict[str, object]:
    """Normalize MCP wrapper return values into the shared response envelope."""

    structured_content = getattr(response, "structuredContent", None)
    if isinstance(structured_content, dict):
        return structured_content
    if isinstance(response, tuple) and len(response) == 2 and isinstance(response[1], dict):
        return response[1]
    raise AssertionError(f"Unsupported tool response type: {type(response)!r}")


@unittest.skipUnless(MCP_AVAILABLE, "The `mcp` package is required for server wrapper tests.")
class ServerToolTests(unittest.TestCase):
    """Verify tool wrappers return the shared envelope for all core tools."""

    def test_open_workbook_success(self) -> None:
        """Verify ``open_workbook`` returns a success envelope on service success.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper behavior.
        """

        payload = {"workbook_id": "wb_001", "path": "/tmp/book.xlsx"}
        with patch.object(server.excel_service, "open_workbook", return_value=payload):
            response = unwrap_tool_response(server.open_workbook("/tmp/book.xlsx"))

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"], payload)

    def test_get_sheet_state_error(self) -> None:
        """Verify service errors are mapped into the shared error envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper error handling.
        """

        with patch.object(
            server.excel_service,
            "get_sheet_state",
            side_effect=ExcelServiceError("missing sheet"),
        ):
            response = unwrap_tool_response(server.get_sheet_state("wb_001", "Missing"))

        self.assertEqual(response["status"], "error")
        self.assertEqual(response["errors"], ["missing sheet"])

    def test_search_cell_success(self) -> None:
        """Verify ``search_cell`` forwards compact match payloads unchanged."""

        payload = {
            "query": "Visible",
            "kind": "text",
            "scope": "workbook",
            "limit": 10,
            "count": 1,
            "truncated": False,
            "matches": ["Data!E2"],
        }
        with patch.object(server.excel_service, "search_cell", return_value=payload):
            response = unwrap_tool_response(server.search_cell("wb_001", "Visible"))

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"]["matches"], ["Data!E2"])

    def test_search_cell_error(self) -> None:
        """Verify ``search_cell`` maps service failures into the shared envelope."""

        with patch.object(
            server.excel_service,
            "search_cell",
            side_effect=ExcelServiceError("`limit` must be a positive integer."),
        ):
            response = unwrap_tool_response(server.search_cell("wb_001", "Visible", limit=0))

        self.assertEqual(response["status"], "error")
        self.assertEqual(response["errors"], ["`limit` must be a positive integer."])

    def test_get_range_success(self) -> None:
        """Verify ``get_range`` forwards range payloads through the shared envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper behavior.
        """

        payload = {"sheet": "Sheet1", "range": "A1:B2", "rows": 2, "columns": 2, "values": []}
        with patch.object(server.excel_service, "get_range", return_value=payload):
            response = unwrap_tool_response(server.get_range("wb_001", "Sheet1", "A1:B2"))

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"]["range"], "A1:B2")

    def test_get_range_disabled_via_env(self) -> None:
        """When ``EXCEL_MCP_DISABLED_TOOLS`` lists ``get_range``, calls fail fast."""

        with patch.dict(os.environ, {"EXCEL_MCP_DISABLED_TOOLS": "get_range, trace_formula"}):
            response = unwrap_tool_response(server.get_range("wb_001", "Sheet1", "A1"))

        self.assertEqual(response["status"], "error")
        self.assertTrue(any("disabled" in err.lower() for err in response["errors"]))

    def test_set_range_success(self) -> None:
        """Verify ``set_range`` exposes service write results through the MCP envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper behavior.
        """

        payload = {"updated_values": True, "updated_formulas": False}
        with patch.object(server.excel_service, "set_range", return_value=payload):
            response = unwrap_tool_response(server.set_range("wb_001", "Sheet1", "A1", values=[["x"]]))

        self.assertEqual(response["status"], "success")
        self.assertTrue(response["data"]["updated_values"])

    def test_recalculate_success(self) -> None:
        """Verify ``recalculate`` exposes service scan results through the MCP envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper behavior.
        """

        payload = {"recalculated": True, "total_formulas": 3, "total_errors": 0}
        with patch.object(server.excel_service, "recalculate", return_value=payload):
            response = unwrap_tool_response(server.recalculate("wb_001", scope="workbook"))

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"]["total_formulas"], 3)

    def test_local_screenshot_success(self) -> None:
        """Verify ``local_screenshot`` exposes screenshot metadata through the MCP envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper behavior.
        """

        payload = {"image_path": "/tmp/out.png"}
        with patch.object(server.excel_service, "local_screenshot", return_value=payload):
            response = unwrap_tool_response(server.local_screenshot("wb_001", "Sheet1", "A1:B2"))

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"]["image_path"], "/tmp/out.png")

    def test_sheet_screenshot_success(self) -> None:
        """Verify ``sheet_screenshot`` exposes image metadata through the MCP envelope."""

        payload = {"image_path": "/tmp/sheet.png"}
        with patch.object(server.excel_service, "sheet_screenshot", return_value=payload):
            response = unwrap_tool_response(server.sheet_screenshot("/tmp/book.xlsx", "Sheet1"))

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"]["image_path"], "/tmp/sheet.png")

    def test_trace_formula_success(self) -> None:
        """Verify ``trace_formula`` exposes trace results through the MCP envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper behavior.
        """

        payload = {
            "sheet": "TraceData",
            "range": "B2",
            "direction": "precedents",
            "max_depth": 1,
            "complete": True,
            "nodes": [
                {"id": "A2", "sheet": "TraceData", "range": "A2"},
                {"id": "B2", "sheet": "TraceData", "range": "B2"},
            ],
            "edges": [{"from": "A2", "to": "B2"}],
        }
        with patch.object(server.excel_service, "trace_formula", return_value=payload):
            response = unwrap_tool_response(
                server.trace_formula("wb_001", "TraceData", "B2", "precedents", max_depth=1)
            )

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"]["max_depth"], 1)
        self.assertEqual(response["data"]["edges"], [{"from": "A2", "to": "B2"}])


if __name__ == "__main__":
    unittest.main()
