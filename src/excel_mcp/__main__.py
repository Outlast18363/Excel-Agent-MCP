"""Process entrypoint for running the Excel MCP server over stdio."""

from __future__ import annotations

import logging
import sys

from .server import mcp_server


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


def main() -> None:
    """Start the FastMCP server using stdio transport.

    Parameters:
        None.

    Returns:
        ``None``. The process blocks while serving MCP requests.
    """

    configure_logging()
    mcp_server.run(transport="stdio")


if __name__ == "__main__":
    main()
