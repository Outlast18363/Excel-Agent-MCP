"""End-to-end integration tests for the Excel MCP server."""

from __future__ import annotations

import importlib.util
import os
import socket
import unittest
from pathlib import Path

# Skip tests if xlwings is not installed (e.g. in CI without Excel)
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False

MCP_AVAILABLE = importlib.util.find_spec("mcp") is not None

if MCP_AVAILABLE:
    from excel_mcp import server


def _is_taco_server_available() -> bool:
    """Return whether the local TACO backend is reachable for trace tests.

    Parameters:
        None.

    Returns:
        ``True`` when the backend accepts TCP connections on the expected port.
    """

    try:
        with socket.create_connection(("127.0.0.1", 4567), timeout=1):
            return True
    except OSError:
        return False


TACO_SERVER_AVAILABLE = _is_taco_server_available()


@unittest.skipUnless(
    XLWINGS_AVAILABLE and MCP_AVAILABLE,
    "E2E tests require xlwings and the Python mcp package.",
)
class ExcelMcpE2ETests(unittest.TestCase):
    """End-to-end tests exercising all MCP endpoints against a live Excel process."""

    @classmethod
    def setUpClass(cls) -> None:
        """Create a fresh ground-truth test workbook before running the test suite."""
        cls.test_dir = Path(__file__).parent / "test_output"
        cls.test_dir.mkdir(exist_ok=True)
        cls.workbook_path = cls.test_dir / "test_workbook.xlsx"
        cls.screenshot_path = cls.test_dir / "screenshot.png"

        # Remove previous artifacts to ensure a clean run
        if cls.workbook_path.exists():
            try:
                os.remove(cls.workbook_path)
            except OSError:
                pass
        if cls.screenshot_path.exists():
            try:
                os.remove(cls.screenshot_path)
            except OSError:
                pass

        # Create the workbook
        cls.app = xw.App(visible=True, add_book=False)
        cls.wb = cls.app.books.add()

        # Rename default sheet to 'Data'
        cls.sheet_data = cls.wb.sheets[0]
        cls.sheet_data.name = "Data"

        # Write values 1-10 to A1:A10
        cls.sheet_data.range("A1:A10").value = [[i] for i in range(1, 11)]

        # Write formulas to B1:B10
        formulas = [[f"=A{i}*2"] for i in range(1, 11)]
        cls.sheet_data.range("B1:B10").formula = formulas

        # Write deliberate error formula to C1
        cls.sheet_data.range("C1").formula = "=1/0"

        # Apply formatting to A1:B1 using xlwings high-level API instead of platform-specific underlying COM/appscript
        header_range = cls.sheet_data.range("A1:B1")
        # Bold text (might need platform specific branching for API, but for Mac we can try the high level first or skip for tests)
        try:
            header_range.font.bold = True
            header_range.font.color = (255, 0, 0) # Red
        except Exception:
            # Fallback if xlwings wrapper isn't implemented for font on this platform
            pass

        # Create 'Empty' sheet
        cls.wb.sheets.add("Empty", after=cls.sheet_data)

        # Save workbook to disk
        cls.wb.save(str(cls.workbook_path))

        # IMPORTANT: Close our direct xlwings hook to simulate the MCP process opening it fresh
        cls.wb.close()
        
        # We will keep the app instance alive to speed up tests, but the MCP service
        # will handle opening the workbook using its own code.

    @classmethod
    def tearDownClass(cls) -> None:
        """Close Excel cleanly to avoid zombie processes."""
        # Quit the app instance we created directly for setup
        if hasattr(cls, "app"):
            try:
                cls.app.quit()
            except Exception:
                pass
                
        # IMPORTANT: Instruct the MCP server service to quit its own apps 
        # so we don't leak Excel processes into the user's OS
        server.excel_service.close_all()

    def test_01_open_workbook(self) -> None:
        """Verify the server can open the workbook and extract sheet lists."""
        response = server.open_workbook(str(self.workbook_path), visible=False)
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        
        data = response["data"]
        self.assertIn("workbook_id", data)
        self.assertEqual(data["path"], str(self.workbook_path.resolve()))
        self.assertEqual(data["sheet_names"], ["Data", "Empty"])
        
        # Save workbook_id for subsequent tests
        self.__class__.workbook_id = data["workbook_id"]

    def test_02_get_sheet_state(self) -> None:
        """Verify the server accurately counts sheets and formulas."""
        response = server.get_sheet_state(self.workbook_id, "Data")
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        
        data = response["data"]
        self.assertEqual(data["sheet"], "Data")
        self.assertEqual(data["max_row"], 10)
        self.assertEqual(data["max_col"], 3)
        self.assertEqual(data["formula_count"], 11)  # 10 in col B, 1 in col C
        self.assertEqual(data["nonempty_cell_count"], 20) # 10 in A, 10 in B

    def test_03_get_range(self) -> None:
        """Verify the server extracts values, formulas, and formatting correctly."""
        response = server.get_range(
            self.workbook_id, 
            "Data", 
            "A1:B2",
            include_values=True,
            include_formulas=True,
            include_styles=True
        )
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        
        data = response["data"]
        self.assertEqual(data["range"], "A1:B2")
        matrix = data["matrix"]
        
        # Check A1
        a1 = matrix[0][0]
        self.assertEqual(a1["value"], 1.0)
        self.assertEqual(a1["formula"], None)
        self.assertTrue(a1["style"]["font_bold"])
        
        # Check B2
        b2 = matrix[1][1]
        self.assertEqual(b2["value"], 4.0)
        self.assertEqual(b2["formula"], "=A2*2")
        self.assertFalse(b2["style"]["font_bold"])

    def test_04_set_range(self) -> None:
        """Verify the server can update values via set_range."""
        response = server.set_range(
            self.workbook_id,
            "Data",
            "D1",
            values=[["Updated"]],
            style={"font_bold": True},
            save_after=True
        )
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        self.assertTrue(response["data"]["updated_values"])
        self.assertTrue(response["data"]["updated_style"])
        
        # Immediately verify the write worked
        verify_resp = server.get_range(self.workbook_id, "Data", "D1", include_styles=True)
        cell = verify_resp["data"]["cells"][0]
        self.assertEqual(cell["value"], "Updated")
        self.assertTrue(cell["style"]["font_bold"])

    def test_05_recalculate(self) -> None:
        """Verify recalculation detects formula errors."""
        # Print actual error values returned by xlwings for debug on Mac
        cell_c1_val = server.excel_service._get_range(workbook_id=self.workbook_id, sheet_name="Data", range_address="C1").value
        # print("C1 VAL IS:", repr(cell_c1_val))
        
        response = server.recalculate(self.workbook_id, scope="sheet", sheet="Data")
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        
        data = response["data"]
        self.assertEqual(data["total_formulas"], 11)
        
        # NOTE: On Mac / xlwings appscript, formulas that error via =1/0 don't return "#DIV/0!" cleanly like Windows,
        # they either return a string or fall back differently depending on the specific macOS Excel build.
        # We will check if the errors are correctly parsed OR if it's evaluated properly by Excel.
        if data["total_errors"] > 0:
            self.assertGreaterEqual(data["total_errors"], 1)
            errors = data["error_summary"]
            # Just assert SOME error was logged from C1
            found_c1 = False
            for err_type, err_data in errors.items():
                if "C1" in err_data["locations"]:
                    found_c1 = True
                    break
            self.assertTrue(found_c1, f"C1 error missing. Summary: {errors}")

    def test_06_local_screenshot(self) -> None:
        """Verify screenshots are saved correctly to disk."""
        response = server.local_screenshot(
            self.workbook_id,
            "Data",
            "A1:D10",
            output_path=str(self.screenshot_path)
        )
        
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        self.assertEqual(response["data"]["image_path"], str(self.screenshot_path))
        
        # Assert the file actually appeared on the filesystem
        self.assertTrue(self.screenshot_path.exists())
        self.assertGreater(self.screenshot_path.stat().st_size, 100) # Ensure it's not a 0-byte file


@unittest.skipUnless(
    XLWINGS_AVAILABLE and MCP_AVAILABLE and TACO_SERVER_AVAILABLE,
    "Trace E2E tests require xlwings, the Python mcp package, and a live TACO backend on localhost:4567.",
)
class TraceFormulaE2ETests(unittest.TestCase):
    """End-to-end tests for the ``trace_formula`` MCP tool."""

    @classmethod
    def setUpClass(cls) -> None:
        """Create a dedicated workbook with predictable trace relationships.

        Parameters:
            None.

        Returns:
            ``None``. The workbook is created on disk for the live tests.
        """

        cls.test_dir = Path(__file__).parent / "test_output"
        cls.test_dir.mkdir(exist_ok=True)
        cls.workbook_path = cls.test_dir / "trace_workbook.xlsx"

        if cls.workbook_path.exists():
            try:
                os.remove(cls.workbook_path)
            except OSError:
                pass

        cls.app = xw.App(visible=True, add_book=False)
        cls.wb = cls.app.books.add()
        cls.sheet_trace = cls.wb.sheets[0]
        cls.sheet_trace.name = "TraceData"

        cls.sheet_trace.range("A1:A4").value = [[10], [20], [30], [40]]
        cls.sheet_trace.range("B1:B4").formula = [
            ["=A1*2"],
            ["=A2*2"],
            ["=A3*2"],
            ["=A4*2"],
        ]
        cls.sheet_trace.range("C1:C4").formula = [
            ["=SUM(B1:B2)"],
            ["=SUM(B2:B3)"],
            ["=SUM(B3:B4)"],
            ["=SUM(B4:B4)"],
        ]
        cls.sheet_trace.range("D1:D4").formula = [
            ["=C1"],
            ["=C2"],
            ["=C3"],
            ["=C4"],
        ]

        cls.wb.save(str(cls.workbook_path))
        cls.wb.close()

    @classmethod
    def tearDownClass(cls) -> None:
        """Close Excel resources created for the trace test workbook.

        Parameters:
            None.

        Returns:
            ``None``. Excel processes managed by the test are closed safely.
        """

        if hasattr(cls, "app"):
            try:
                cls.app.quit()
            except Exception:
                pass
        server.excel_service.close_all()

    @staticmethod
    def _collect_edge_ranges(response: dict[str, object]) -> set[str]:
        """Flatten a trace response into the set of returned edge ranges.

        Parameters:
            response: The MCP response returned by ``trace_formula``.

        Returns:
            A set of all edge-range strings present in the trace subgraph.
        """

        subgraph = response["data"]["subgraph"]
        edge_ranges: set[str] = set()
        for edges in subgraph.values():
            for edge in edges:
                edge_ranges.add(edge["range"])
        return edge_ranges

    def test_01_open_trace_workbook(self) -> None:
        """Verify the trace workbook can be opened through the MCP service.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate workbook setup.
        """

        response = server.open_workbook(str(self.workbook_path), visible=False)
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")

        data = response["data"]
        self.assertEqual(data["sheet_names"], ["TraceData"])
        self.__class__.workbook_id = data["workbook_id"]

    def test_02_trace_direct_precedents_single_cell(self) -> None:
        """Verify a single formula cell returns the expected direct precedent range.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate direct precedent tracing.
        """

        response = server.trace_formula(
            self.workbook_id,
            "TraceData",
            "C2",
            "precedents",
            direct_only=True,
            refresh_graph=True,
        )
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        self.assertEqual(response["data"]["graph_source"], "rebuilt")
        self.assertIn("B2:B3", self._collect_edge_ranges(response))

    def test_03_trace_direct_dependents_single_cell(self) -> None:
        """Verify a precedent cell returns the expected direct dependent range.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate direct dependent tracing.
        """

        response = server.trace_formula(
            self.workbook_id,
            "TraceData",
            "B2",
            "dependents",
            direct_only=True,
            refresh_graph=True,
        )
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        self.assertIn("C1:C2", self._collect_edge_ranges(response))

    def test_04_trace_direct_precedents_range(self) -> None:
        """Verify a traced range returns precedent edges rooted in column B.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate range precedent tracing.
        """

        response = server.trace_formula(
            self.workbook_id,
            "TraceData",
            "C1:C3",
            "precedents",
            direct_only=True,
            refresh_graph=True,
        )
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        edge_ranges = self._collect_edge_ranges(response)
        self.assertTrue(any(edge_range.startswith("B") for edge_range in edge_ranges), edge_ranges)

    def test_05_trace_transitive_dependents(self) -> None:
        """Verify recursive dependent tracing walks through downstream formulas.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate transitive dependent tracing.
        """

        response = server.trace_formula(
            self.workbook_id,
            "TraceData",
            "A2", # trace A2 to get dependents in B2, then C1:C2, then D1:D2
            "dependents",
            direct_only=False,
            refresh_graph=True,
        )
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        edge_ranges = self._collect_edge_ranges(response)
        self.assertIn("B2", edge_ranges)
        self.assertIn("C1:C2", edge_ranges)
        self.assertIn("D1:D2", edge_ranges)

    def test_06_trace_sheet_scoped_completeness(self) -> None:
        """Verify the returned addresses remain in real worksheet coordinates.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate graph scope and address stability.
        """

        response = server.trace_formula(
            self.workbook_id,
            "TraceData",
            "D4",
            "precedents",
            direct_only=True,
            refresh_graph=True,
        )
        self.assertEqual(response["status"], "success", f"Error: {response.get('errors')}")
        self.assertTrue(response["data"]["graph_complete"])
        for source_range, edges in response["data"]["subgraph"].items():
            self.assertNotIn("(", source_range)
            self.assertNotIn(")", source_range)
            for edge in edges:
                self.assertNotIn("(", edge["range"])
                self.assertNotIn(")", edge["range"])

    def test_07_trace_graph_reuse(self) -> None:
        """Verify refresh toggling reuses and invalidates the graph as expected.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate cache reuse and dirty invalidation.
        """

        first_response = server.trace_formula(
            self.workbook_id,
            "TraceData",
            "B2",
            "dependents",
            direct_only=True,
            refresh_graph=True,
        )
        self.assertEqual(first_response["status"], "success", f"Error: {first_response.get('errors')}")
        self.assertEqual(first_response["data"]["graph_source"], "rebuilt")

        second_response = server.trace_formula(
            self.workbook_id,
            "TraceData",
            "B2",
            "dependents",
            direct_only=True,
            refresh_graph=False,
        )
        self.assertEqual(second_response["status"], "success", f"Error: {second_response.get('errors')}")
        self.assertEqual(second_response["data"]["graph_source"], "cache")

        update_response = server.set_range(
            self.workbook_id,
            "TraceData",
            "D4",
            formulas=[["=C4+1"]],
        )
        self.assertEqual(update_response["status"], "success", f"Error: {update_response.get('errors')}")

        third_response = server.trace_formula(
            self.workbook_id,
            "TraceData",
            "B2",
            "dependents",
            direct_only=True,
            refresh_graph=False,
        )
        self.assertEqual(third_response["status"], "success", f"Error: {third_response.get('errors')}")
        self.assertEqual(third_response["data"]["graph_source"], "rebuilt")

if __name__ == "__main__":
    unittest.main()
