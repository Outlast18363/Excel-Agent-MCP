"""Unit tests for pure helper logic in the Excel MCP server."""

from __future__ import annotations

import unittest
from datetime import date, datetime

from excel_mcp.helpers import (
    ExcelServiceError,
    column_number_to_name,
    hex_to_rgb_tuple,
    normalize_formula_grid,
    normalize_matrix_input,
    normalize_taco_pattern,
    normalize_taco_ref_key,
    taco_ref_to_a1_address,
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

    def test_taco_ref_to_a1_address_converts_zero_based_bounds(self) -> None:
        """Verify TACO ref payloads convert into worksheet A1 ranges.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate ref serialization.
        """

        self.assertEqual(
            taco_ref_to_a1_address(
                {"_row": 1, "_column": 1, "_lastRow": 2, "_lastColumn": 2}
            ),
            "B2:C3",
        )

    def test_normalize_taco_ref_key_strips_range_parentheses(self) -> None:
        """Verify serialized TACO map keys normalize into plain A1 text.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate key cleanup.
        """

        self.assertEqual(normalize_taco_ref_key("(A1:B2)"), "A1:B2")
        self.assertEqual(normalize_taco_ref_key("C3"), "C3")

    def test_normalize_taco_pattern_maps_internal_enum_names(self) -> None:
        """Verify internal TACO enum names map into stable public labels.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate pattern-name normalization.
        """

        self.assertEqual(normalize_taco_pattern("TYPEONE"), "RR")
        self.assertEqual(normalize_taco_pattern("NOTYPE"), "NO_COMP")

    def test_normalize_formula_grid_preserves_shape_and_blanks(self) -> None:
        """Verify xlwings-style formula outputs normalize into dense matrices.

        Parameters:
            None.

        Returns:
            ``None``. Assertions validate formula-grid normalization.
        """

        self.assertEqual(
            normalize_formula_grid([["=A1*2", None], [3, "text"]], 2, 2),
            [["=A1*2", ""], ["3", "text"]],
        )


if __name__ == "__main__":
    unittest.main()
