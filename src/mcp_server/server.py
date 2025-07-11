#!/usr/bin/env python3
"""
MCP Sheet Parser Server

A Model Context Protocol server that provides spreadsheet parsing and HTML conversion tools.
Supports multiple file formats: .xlsx, .xls, .xlsb, .xlsm, .et, .ett, .ets, .csv
"""

import asyncio
import logging
from typing import Any, Dict, Optional

from mcp.server.fastmcp import FastMCP
from src.services.parsing_service import ParsingService, ParsingServiceError


# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create MCP server instance
mcp = FastMCP("Sheet Parser Server")

# Initialize parsing service
parsing_service = ParsingService()


@mcp.tool()
def convert_file_to_html(
    file_path: str,
    sheet_name: Optional[str] = None,
    include_styles: bool = True,
    include_metadata: bool = True
) -> Dict[str, Any]:
    """
    Convert a spreadsheet file to styled HTML table.
    
    This is the convenience tool for one-shot conversion tasks.
    
    Args:
        file_path: Path to the spreadsheet file
        sheet_name: Name of the sheet to convert (optional, defaults to first sheet)
        include_styles: Whether to include styling information (default: True)
        include_metadata: Whether to include metadata section (default: True)
    
    Returns:
        Dictionary containing HTML content and metadata
    """
    try:
        logger.info(f"Converting file to HTML: {file_path}")
        result = parsing_service.convert_to_html(
            file_path=file_path,
            sheet_name=sheet_name,
            include_styles=include_styles,
            include_metadata=include_metadata
        )
        logger.info(f"Successfully converted {file_path} to HTML")
        return result
    
    except ParsingServiceError as e:
        logger.error(f"Parsing service error: {e}")
        return {
            "status": "error",
            "error_type": "parsing_error",
            "error_message": str(e),
            "file_path": file_path
        }
    
    except Exception as e:
        logger.error(f"Unexpected error converting {file_path}: {e}")
        return {
            "status": "error",
            "error_type": "unexpected_error",
            "error_message": f"An unexpected error occurred: {str(e)}",
            "file_path": file_path
        }


@mcp.tool()
def get_sheet_metadata(file_path: str) -> Dict[str, Any]:
    """
    Get metadata about a spreadsheet file without full parsing.
    
    This is a professional tool for quick file inspection.
    
    Args:
        file_path: Path to the spreadsheet file
    
    Returns:
        Dictionary containing file and sheet metadata
    """
    try:
        logger.info(f"Getting metadata for: {file_path}")
        result = parsing_service.get_sheet_metadata(file_path)
        logger.info(f"Successfully retrieved metadata for {file_path}")
        return result
    
    except ParsingServiceError as e:
        logger.error(f"Parsing service error: {e}")
        return {
            "status": "error",
            "error_type": "parsing_error",
            "error_message": str(e),
            "file_path": file_path
        }
    
    except Exception as e:
        logger.error(f"Unexpected error getting metadata for {file_path}: {e}")
        return {
            "status": "error",
            "error_type": "unexpected_error",
            "error_message": f"An unexpected error occurred: {str(e)}",
            "file_path": file_path
        }


@mcp.tool()
def parse_sheet_to_json(
    file_path: str,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """
    Parse a spreadsheet sheet and return structured JSON data.
    
    This is a professional tool for data extraction workflows.
    
    Args:
        file_path: Path to the spreadsheet file
        sheet_name: Name of the sheet to parse (optional, defaults to first sheet)
    
    Returns:
        Dictionary containing structured sheet data in JSON format
    """
    try:
        logger.info(f"Parsing sheet to JSON: {file_path}")
        result = parsing_service.parse_sheet_to_json(
            file_path=file_path,
            sheet_name=sheet_name
        )
        logger.info(f"Successfully parsed {file_path} to JSON")
        return result
    
    except ParsingServiceError as e:
        logger.error(f"Parsing service error: {e}")
        return {
            "status": "error",
            "error_type": "parsing_error",
            "error_message": str(e),
            "file_path": file_path
        }
    
    except Exception as e:
        logger.error(f"Unexpected error parsing {file_path}: {e}")
        return {
            "status": "error",
            "error_type": "unexpected_error",
            "error_message": f"An unexpected error occurred: {str(e)}",
            "file_path": file_path
        }


@mcp.tool()
def convert_json_to_html(
    sheet_json: Dict[str, Any],
    include_styles: bool = True,
    include_metadata: bool = True
) -> Dict[str, Any]:
    """
    Convert JSON sheet data to styled HTML table.
    
    This is a professional tool for multi-step workflows where data has been
    previously parsed to JSON format.
    
    Args:
        sheet_json: JSON representation of sheet data (from parse_sheet_to_json)
        include_styles: Whether to include styling information (default: True)
        include_metadata: Whether to include metadata section (default: True)
    
    Returns:
        Dictionary containing HTML content and metadata
    """
    try:
        logger.info("Converting JSON data to HTML")
        result = parsing_service.convert_json_to_html(
            sheet_json=sheet_json,
            include_styles=include_styles,
            include_metadata=include_metadata
        )
        logger.info("Successfully converted JSON to HTML")
        return result
    
    except ParsingServiceError as e:
        logger.error(f"Parsing service error: {e}")
        return {
            "status": "error",
            "error_type": "parsing_error",
            "error_message": str(e)
        }
    
    except Exception as e:
        logger.error(f"Unexpected error converting JSON to HTML: {e}")
        return {
            "status": "error",
            "error_type": "unexpected_error",
            "error_message": f"An unexpected error occurred: {str(e)}"
        }


def main():
    """Main entry point for the MCP server."""
    logger.info("Starting MCP Sheet Parser Server...")
    logger.info("Supported formats: .xlsx, .xls, .xlsb, .xlsm, .et, .ett, .ets, .csv")
    logger.info("Available tools: convert_file_to_html, get_sheet_metadata, parse_sheet_to_json, convert_json_to_html")
    
    try:
        # Run the MCP server
        mcp.run()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server error: {e}")
        raise


if __name__ == "__main__":
    main()
