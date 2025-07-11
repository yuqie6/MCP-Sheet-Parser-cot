"""
MCP Tools Registration

This module defines and registers all MCP tools for the Sheet Parser server.
"""

import logging
from typing import Any, Dict, List, Optional
from pathlib import Path

from mcp.server import Server
from mcp.types import Tool, TextContent

from ..services.sheet_service import SheetService
from ..parsers.factory import ParserFactory
from ..converters.html_converter import HTMLConverter

logger = logging.getLogger(__name__)


def register_tools(server: Server) -> None:
    """Register all MCP tools with the server."""
    
    # Initialize core services
    parser_factory = ParserFactory()
    html_converter = HTMLConverter()
    sheet_service = SheetService(parser_factory, html_converter)
    
    @server.list_tools()
    async def handle_list_tools() -> List[Tool]:
        """List all available tools."""
        return [
            Tool(
                name="parse_sheet_to_json",
                description="Parse a spreadsheet file and return structured JSON data for LLM analysis",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to parse"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="convert_json_to_html",
                description="Convert JSON spreadsheet data to a perfect HTML file with 95% style fidelity",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "json_data": {
                            "type": "object",
                            "description": "JSON data from parse_sheet_to_json"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the HTML file"
                        }
                    },
                    "required": ["json_data", "output_path"]
                }
            ),
            Tool(
                name="convert_file_to_html",
                description="Smart direct conversion from file to HTML with intelligent processing strategy",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to convert"
                        },
                        "compact_mode": {
                            "type": "boolean",
                            "description": "Enable compact HTML mode for size optimization",
                            "default": True
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="convert_file_to_html_file",
                description="Convert spreadsheet file directly to HTML file (large file friendly)",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to convert"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the HTML file"
                        }
                    },
                    "required": ["file_path", "output_path"]
                }
            ),
            Tool(
                name="get_table_summary",
                description="Get a quick preview and summary of spreadsheet content",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to analyze"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="get_sheet_metadata",
                description="Get detailed metadata about the spreadsheet file",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to analyze"
                        }
                    },
                    "required": ["file_path"]
                }
            )
        ]
    
    @server.call_tool()
    async def handle_call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent]:
        """Handle tool calls."""
        try:
            if name == "parse_sheet_to_json":
                return await _handle_parse_sheet_to_json(arguments, sheet_service)
            elif name == "convert_json_to_html":
                return await _handle_convert_json_to_html(arguments, sheet_service)
            elif name == "convert_file_to_html":
                return await _handle_convert_file_to_html(arguments, sheet_service)
            elif name == "convert_file_to_html_file":
                return await _handle_convert_file_to_html_file(arguments, sheet_service)
            elif name == "get_table_summary":
                return await _handle_get_table_summary(arguments, sheet_service)
            elif name == "get_sheet_metadata":
                return await _handle_get_sheet_metadata(arguments, sheet_service)
            else:
                raise ValueError(f"Unknown tool: {name}")
                
        except Exception as e:
            logger.error(f"Error in tool {name}: {e}")
            return [TextContent(
                type="text",
                text=f"Error: {str(e)}"
            )]


# Tool handler functions (to be implemented in subsequent tasks)
async def _handle_parse_sheet_to_json(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle parse_sheet_to_json tool call."""
    # TODO: Implement in Task 3 (JSON converter)
    return [TextContent(
        type="text",
        text="parse_sheet_to_json: Not yet implemented - will be completed in Task 3"
    )]


async def _handle_convert_json_to_html(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle convert_json_to_html tool call."""
    # TODO: Implement in Task 3 (JSON converter)
    return [TextContent(
        type="text",
        text="convert_json_to_html: Not yet implemented - will be completed in Task 3"
    )]


async def _handle_convert_file_to_html(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle convert_file_to_html tool call."""
    # Basic implementation using existing service
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("file_path is required")
    
    if not Path(file_path).exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    
    try:
        html_content = service.convert_to_html(file_path)
        return [TextContent(
            type="text",
            text=f"HTML conversion completed. Content length: {len(html_content)} characters.\n\n{html_content}"
        )]
    except Exception as e:
        raise RuntimeError(f"Failed to convert file to HTML: {str(e)}")


async def _handle_convert_file_to_html_file(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle convert_file_to_html_file tool call."""
    # TODO: Implement file output functionality
    return [TextContent(
        type="text",
        text="convert_file_to_html_file: Not yet implemented - will be completed in Task 5"
    )]


async def _handle_get_table_summary(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle get_table_summary tool call."""
    # TODO: Implement summary functionality
    return [TextContent(
        type="text",
        text="get_table_summary: Not yet implemented - will be completed in Task 5"
    )]


async def _handle_get_sheet_metadata(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle get_sheet_metadata tool call."""
    # TODO: Implement metadata functionality
    return [TextContent(
        type="text",
        text="get_sheet_metadata: Not yet implemented - will be completed in Task 5"
    )]
