#!/usr/bin/env python3
"""
MCP Sheet Parser Server - Main Entry Point

A Model Context Protocol server that provides spreadsheet parsing and HTML conversion tools.
Supports multiple file formats and provides both convenience and professional tool interfaces.
"""

import asyncio
from src.mcp_server.server import main as server_main


def main():
    """Main entry point for the MCP Sheet Parser Server."""
    asyncio.run(server_main())


if __name__ == "__main__":
    main()
