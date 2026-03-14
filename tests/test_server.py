"""Unit tests for the MCP tool wrappers exposed by the Excel server."""

from __future__ import annotations

import importlib.util
import unittest
from unittest.mock import patch

from excel_mcp.helpers import ExcelServiceError

MCP_AVAILABLE = importlib.util.find_spec("mcp") is not None

if MCP_AVAILABLE:
    from excel_mcp import server


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
            response = server.open_workbook("/tmp/book.xlsx")

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
            response = server.get_sheet_state("wb_001", "Missing")

        self.assertEqual(response["status"], "error")
        self.assertEqual(response["errors"], ["missing sheet"])

    def test_get_range_success(self) -> None:
        """Verify ``get_range`` forwards range payloads through the shared envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper behavior.
        """

        payload = {"sheet": "Sheet1", "range": "A1:B2", "matrix": [], "cells": []}
        with patch.object(server.excel_service, "get_range", return_value=payload):
            response = server.get_range("wb_001", "Sheet1", "A1:B2")

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"]["range"], "A1:B2")

    def test_set_range_success(self) -> None:
        """Verify ``set_range`` exposes service write results through the MCP envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate wrapper behavior.
        """

        payload = {"updated_values": True, "updated_formulas": False}
        with patch.object(server.excel_service, "set_range", return_value=payload):
            response = server.set_range("wb_001", "Sheet1", "A1", values=[["x"]])

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
            response = server.recalculate("wb_001", scope="workbook")

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
            response = server.local_screenshot("wb_001", "Sheet1", "A1:B2")

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"]["image_path"], "/tmp/out.png")


if __name__ == "__main__":
    unittest.main()
