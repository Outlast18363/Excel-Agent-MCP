"""Shared response and JSON-normalization helpers for the Excel MCP server."""

from __future__ import annotations

from collections.abc import Sequence
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any, TypeAlias, TypedDict


JsonPrimitive: TypeAlias = str | int | float | bool | None
JsonValue: TypeAlias = JsonPrimitive | list[Any] | dict[str, Any]

class McpResponse(TypedDict):
    """Typed response envelope used by all MCP tools."""

    status: str
    data: JsonValue
    warnings: list[str]
    errors: list[str]


def normalize_excel_value(value: Any) -> JsonValue:
    """Convert an Excel value into JSON-safe primitives.

    Parameters:
        value: Any value returned by Excel or Python helper code.

    Returns:
        A JSON-safe value that preserves nested shape where possible.
    """

    if value is None:
        return None

    if isinstance(value, (str, int, float, bool)):
        return value

    if isinstance(value, Decimal):
        return float(value)

    if isinstance(value, datetime):
        return value.isoformat()

    if isinstance(value, date):
        return value.isoformat()

    if isinstance(value, time):
        return value.isoformat()

    if isinstance(value, dict):
        return {
            str(key): normalize_excel_value(item)
            for key, item in value.items()
        }

    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray)):
        return [normalize_excel_value(item) for item in value]

    return str(value)


def make_response(
    *,
    status: str,
    data: JsonValue | None = None,
    warnings: list[str] | None = None,
    errors: list[str] | None = None,
) -> McpResponse:
    """Build the shared MCP response envelope.

    Parameters:
        status: The high-level outcome, usually ``success`` or ``error``.
        data: The tool payload to return to the MCP client.
        warnings: Optional non-fatal warning messages.
        errors: Optional fatal or user-facing error messages.

    Returns:
        A dictionary matching the stable response contract used by all tools.
    """

    return McpResponse(
        status=status,
        data=normalize_excel_value(data),
        warnings=warnings or [],
        errors=errors or [],
    )


def success_response(
    data: JsonValue | None = None,
    warnings: list[str] | None = None,
) -> McpResponse:
    """Build a success response with the shared envelope.

    Parameters:
        data: The payload to include in the response.
        warnings: Optional warning messages to surface to the caller.

    Returns:
        A success response dictionary.
    """

    return make_response(status="success", data=data, warnings=warnings, errors=[])


def error_response(
    message: str,
    *,
    data: JsonValue | None = None,
    warnings: list[str] | None = None,
) -> McpResponse:
    """Build an error response with a single primary message.

    Parameters:
        message: The main error string to return to the caller.
        data: Optional partial payload to include with the error.
        warnings: Optional warning messages to include.

    Returns:
        An error response dictionary with the shared envelope shape.
    """

    return make_response(status="error", data=data, warnings=warnings, errors=[message])
