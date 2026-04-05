"""Process entrypoint for running the Excel MCP server over stdio."""

from __future__ import annotations

import logging
import sys

from .server import mcp_server
from .service import excel_service


def configure_logging() -> None:
    """Send server logs to stderr so stdout stays reserved for MCP traffic.

    Parameters:
        None.

    Returns:
        ``None``. The root logger is configured in place.
    """

    logging.basicConfig(
        level=logging.INFO,
        stream=sys.stderr,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


def shutdown_excel_service() -> None:
    """Close managed Excel workbooks and apps during server shutdown.

    Parameters:
        None.

    Returns:
        ``None``. The cleanup runs on a best-effort basis and swallows
        shutdown-time exceptions so process exit is not blocked.
    """

    try:
        excel_service.close_all()
    except Exception:  # pragma: no cover - shutdown cleanup must stay best-effort
        logging.exception("Failed to close Excel resources during shutdown.")


def main() -> None:
    """Start the FastMCP server using stdio transport.

    Parameters:
        None.

    Returns:
        ``None``. The process blocks while serving MCP requests.
    """

    configure_logging()
    try:
        mcp_server.run(transport="stdio")
    finally:
        shutdown_excel_service()


if __name__ == "__main__":
    main()
