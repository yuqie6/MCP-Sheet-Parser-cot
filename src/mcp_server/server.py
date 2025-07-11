#!/usr/bin/env python3
"""
MCP Sheet Parser Server

A Model Context Protocol server that provides spreadsheet parsing and HTML conversion tools.
Supports multiple file formats and provides both convenience and professional tool interfaces.
"""

import asyncio
import logging
from typing import Any, Dict

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool

from .tools import register_tools

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def create_server() -> Server:
    """Create and configure the MCP server."""
    server = Server("mcp-sheet-parser")
    
    # Register all tools
    register_tools(server)
    
    logger.info("MCP Sheet Parser Server initialized")
    return server


async def main():
    """Main entry point for the MCP Sheet Parser Server."""
    try:
        server = create_server()
        
        # Run the server using stdio transport
        async with stdio_server() as (read_stream, write_stream):
            logger.info("Starting MCP Sheet Parser Server...")
            await server.run(
                read_stream,
                write_stream,
                server.create_initialization_options()
            )
    except Exception as e:
        logger.error(f"Server error: {e}")
        raise


if __name__ == "__main__":
    asyncio.run(main())
