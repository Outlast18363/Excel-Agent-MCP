"""Unit tests for pure helper logic in the Excel MCP server."""

from __future__ import annotations

import unittest
from datetime import date, datetime

from excel_mcp.helpers import (
    ExcelServiceError,
    build_trace_node_payload,
    column_number_to_name,
    expand_formulas_ref,
    format_formulas_ref,
    hex_to_rgb_tuple,
    normalize_formula_grid,
    normalize_matrix_input,
    normalize_number_format_grid,
    normalize_range_read_matrix,
    normalize_trace_ref,
    parse_formulas_ref,
    style_payload_key,
    validate_matrix_shape,
)
from excel_mcp.types import error_response, normalize_excel_value, success_response


class HelperTests(unittest.TestCase):
    """Exercise helper functions that do not require a live Excel process."""

    def test_normalize_excel_value_serializes_dates(self) -> None:
        """Verify date-like values are converted into JSON-safe ISO strings.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate normalization behavior.
        """

        value = {
            "created": datetime(2026, 3, 13, 9, 30, 0),
            "due": date(2026, 3, 14),
            "items": [1, None, "ok"],
        }

        normalized = normalize_excel_value(value)

        self.assertEqual(normalized["created"], "2026-03-13T09:30:00")
        self.assertEqual(normalized["due"], "2026-03-14")
        self.assertEqual(normalized["items"], [1, None, "ok"])

    def test_success_response_uses_shared_envelope(self) -> None:
        """Verify success responses preserve the stable MCP envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate response shape.
        """

        response = success_response({"sheet": "Summary"})

        self.assertEqual(response["status"], "success")
        self.assertEqual(response["data"], {"sheet": "Summary"})
        self.assertEqual(response["warnings"], [])
        self.assertEqual(response["errors"], [])

    def test_error_response_uses_shared_envelope(self) -> None:
        """Verify error responses preserve the stable MCP envelope.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate response shape.
        """

        response = error_response("boom")

        self.assertEqual(response["status"], "error")
        self.assertEqual(response["data"], None)
        self.assertEqual(response["warnings"], [])
        self.assertEqual(response["errors"], ["boom"])

    def test_column_number_to_name_supports_multi_letter_columns(self) -> None:
        """Verify numeric Excel columns convert into the expected labels.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate column-name conversion.
        """

        self.assertEqual(column_number_to_name(1), "A")
        self.assertEqual(column_number_to_name(26), "Z")
        self.assertEqual(column_number_to_name(27), "AA")
        self.assertEqual(column_number_to_name(52), "AZ")

    def test_hex_to_rgb_tuple_parses_hex_colors(self) -> None:
        """Verify hex color strings convert into RGB tuples.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate RGB conversion.
        """

        self.assertEqual(hex_to_rgb_tuple("#123ABC"), (18, 58, 188))

    def test_normalize_matrix_input_wraps_single_cell_scalar(self) -> None:
        """Verify a scalar write payload is accepted for a single-cell target.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate matrix normalization.
        """

        self.assertEqual(normalize_matrix_input("ok", 1, 1, "values"), [["ok"]])

    def test_validate_matrix_shape_rejects_invalid_shapes(self) -> None:
        """Verify invalid matrix shapes raise clear service errors.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate shape checking.
        """

        with self.assertRaises(ExcelServiceError):
            validate_matrix_shape([[1, 2]], 2, 2, "values")

    def test_parse_formulas_ref_splits_workbook_sheet_and_range(self) -> None:
        """Verify formulas refs split into workbook, sheet, and range parts.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate formulas-ref parsing.
        """

        self.assertEqual(
            parse_formulas_ref("'[book.xlsx]SHEET ONE'!B2:C3"),
            ("book.xlsx", "SHEET ONE", "B2:C3"),
        )

    def test_format_formulas_ref_builds_qualified_ref(self) -> None:
        """Verify formulas refs can be rebuilt from workbook, sheet, and range.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate formulas-ref formatting.
        """

        self.assertEqual(
            format_formulas_ref("book.xlsx", "SHEET1", "A1:B2"),
            "'[book.xlsx]SHEET1'!A1:B2",
        )

    def test_expand_formulas_ref_expands_ranges_into_single_cells(self) -> None:
        """Verify formulas range refs expand into workbook-qualified cell refs.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate range expansion.
        """

        self.assertEqual(
            expand_formulas_ref("'[book.xlsx]INPUTS'!A1:B2"),
            [
                "'[book.xlsx]INPUTS'!A1",
                "'[book.xlsx]INPUTS'!B1",
                "'[book.xlsx]INPUTS'!A2",
                "'[book.xlsx]INPUTS'!B2",
            ],
        )

    def test_normalize_trace_ref_omits_same_sheet_prefix(self) -> None:
        """Verify same-sheet trace refs normalize into plain A1 addresses.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate same-sheet normalization.
        """

        self.assertEqual(
            normalize_trace_ref(
                "'[book.xlsx]CALC'!B2:C3",
                "Calc",
                {"CALC": "Calc", "INPUTS": "Inputs"},
            ),
            "B2:C3",
        )

    def test_normalize_trace_ref_keeps_cross_sheet_prefix(self) -> None:
        """Verify cross-sheet trace refs keep the display sheet prefix.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate cross-sheet normalization.
        """

        self.assertEqual(
            normalize_trace_ref(
                "'[book.xlsx]INPUTS'!A1",
                "Calc",
                {"CALC": "Calc", "INPUTS": "Inputs"},
            ),
            "Inputs!A1",
        )

    def test_build_trace_node_payload_includes_address_parts(self) -> None:
        """Verify node payloads include address metadata when requested.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate node-payload formatting.
        """

        self.assertEqual(
            build_trace_node_payload(
                "'[book.xlsx]INPUTS'!A1:A2",
                "Calc",
                {"CALC": "Calc", "INPUTS": "Inputs"},
                True,
            ),
            {"id": "Inputs!A1:A2", "sheet": "Inputs", "range": "A1:A2"},
        )

    def test_build_trace_node_payload_can_skip_split_metadata(self) -> None:
        """Verify node payloads can omit split address metadata.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate compact node-payload formatting.
        """

        self.assertEqual(
            build_trace_node_payload(
                "'[book.xlsx]CALC'!C1",
                "Calc",
                {"CALC": "Calc"},
                False,
            ),
            {"id": "C1"},
        )

    def test_normalize_formula_grid_preserves_shape_and_blanks(self) -> None:
        """Verify xlwings-style formula outputs normalize into dense matrices.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate formula-grid normalization.
        """

        self.assertEqual(
            normalize_formula_grid([["=A1*2", None], [3, "text"]], 2, 2),
            [["=A1*2", None], [None, None]],
        )

    def test_normalize_range_read_matrix_supports_single_row_outputs(self) -> None:
        """Verify one-row xlwings reads become a single nested list.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate read-matrix normalization.
        """

        self.assertEqual(
            normalize_range_read_matrix([1, 2, 3], 1, 3, "values"),
            [[1, 2, 3]],
        )

    def test_normalize_range_read_matrix_supports_single_column_outputs(self) -> None:
        """Verify one-column xlwings reads become one value per nested row.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate read-matrix normalization.
        """

        self.assertEqual(
            normalize_range_read_matrix([1, 2, 3], 3, 1, "values"),
            [[1], [2], [3]],
        )

    def test_normalize_number_format_grid_broadcasts_uniform_scalars(self) -> None:
        """Verify one uniform number format can fill an entire target matrix.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate number-format normalization.
        """

        self.assertEqual(
            normalize_number_format_grid("0.00%", 2, 2),
            [["0.00%", "0.00%"], ["0.00%", "0.00%"]],
        )

    def test_style_payload_key_is_stable_for_identical_styles(self) -> None:
        """Verify identical style payloads produce the same deduplication key.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate style-key stability.
        """

        shared_style = {
            "font_name": "Calibri",
            "font_size": 11.0,
            "font_bold": False,
            "font_italic": False,
            "font_color": None,
            "horizontal_alignment": None,
            "vertical_alignment": None,
            "wrap_text": False,
            "fill_color": "#FFFFFF",
        }

        self.assertEqual(style_payload_key(shared_style), style_payload_key(dict(shared_style)))


if __name__ == "__main__":
    unittest.main()
