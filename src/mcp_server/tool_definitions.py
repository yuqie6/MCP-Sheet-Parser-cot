"""
JSON Schema definitions for MCP Sheet Parser Server tools.

This module contains the formal tool definitions and schemas that describe
the interface contract for each MCP tool.
"""

from typing import Dict, Any

# Tool definitions following MCP protocol standards
TOOL_DEFINITIONS = {
    "convert_file_to_html": {
        "name": "convert_file_to_html",
        "description": "Convert a spreadsheet file to styled HTML table. This is the convenience tool for one-shot conversion tasks.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the spreadsheet file. Supported formats: .xlsx, .xls, .xlsb, .xlsm, .et, .ett, .ets, .csv"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Name of the sheet to convert (optional, defaults to first sheet)"
                },
                "include_styles": {
                    "type": "boolean",
                    "description": "Whether to include styling information (default: True)",
                    "default": True
                },
                "include_metadata": {
                    "type": "boolean",
                    "description": "Whether to include metadata section (default: True)",
                    "default": True
                }
            },
            "required": ["file_path"]
        },
        "outputSchema": {
            "type": "object",
            "properties": {
                "status": {
                    "type": "string",
                    "enum": ["success", "error"],
                    "description": "Operation status"
                },
                "html_content": {
                    "type": "string",
                    "description": "Generated HTML content (present when status is success)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Name of the converted sheet"
                },
                "metadata": {
                    "type": "object",
                    "description": "File and conversion metadata"
                },
                "error_type": {
                    "type": "string",
                    "description": "Type of error (present when status is error)"
                },
                "error_message": {
                    "type": "string",
                    "description": "Error description (present when status is error)"
                }
            },
            "required": ["status"]
        }
    },
    
    "get_sheet_metadata": {
        "name": "get_sheet_metadata",
        "description": "Get metadata about a spreadsheet file without full parsing. This is a professional tool for quick file inspection.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the spreadsheet file"
                }
            },
            "required": ["file_path"]
        },
        "outputSchema": {
            "type": "object",
            "properties": {
                "status": {
                    "type": "string",
                    "enum": ["success", "error"],
                    "description": "Operation status"
                },
                "file_path": {
                    "type": "string",
                    "description": "Path to the analyzed file"
                },
                "file_format": {
                    "type": "string",
                    "description": "File format (extension without dot)"
                },
                "file_size": {
                    "type": "integer",
                    "description": "File size in bytes"
                },
                "last_modified": {
                    "type": "number",
                    "description": "Last modification timestamp"
                },
                "sheet_names": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of sheet names in the file"
                },
                "total_sheets": {
                    "type": "integer",
                    "description": "Total number of sheets"
                },
                "error_type": {
                    "type": "string",
                    "description": "Type of error (present when status is error)"
                },
                "error_message": {
                    "type": "string",
                    "description": "Error description (present when status is error)"
                }
            },
            "required": ["status"]
        }
    },
    
    "parse_sheet_to_json": {
        "name": "parse_sheet_to_json",
        "description": "Parse a spreadsheet sheet and return structured JSON data. This is a professional tool for data extraction workflows.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the spreadsheet file"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Name of the sheet to parse (optional, defaults to first sheet)"
                }
            },
            "required": ["file_path"]
        },
        "outputSchema": {
            "type": "object",
            "properties": {
                "status": {
                    "type": "string",
                    "enum": ["success", "error"],
                    "description": "Operation status"
                },
                "sheet_data": {
                    "type": "object",
                    "description": "Structured sheet data in JSON format"
                },
                "metadata": {
                    "type": "object",
                    "description": "Parsing metadata"
                },
                "error_type": {
                    "type": "string",
                    "description": "Type of error (present when status is error)"
                },
                "error_message": {
                    "type": "string",
                    "description": "Error description (present when status is error)"
                }
            },
            "required": ["status"]
        }
    },
    
    "convert_json_to_html": {
        "name": "convert_json_to_html",
        "description": "Convert JSON sheet data to styled HTML table. This is a professional tool for multi-step workflows where data has been previously parsed to JSON format.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "sheet_json": {
                    "type": "object",
                    "description": "JSON representation of sheet data (from parse_sheet_to_json)"
                },
                "include_styles": {
                    "type": "boolean",
                    "description": "Whether to include styling information (default: True)",
                    "default": True
                },
                "include_metadata": {
                    "type": "boolean",
                    "description": "Whether to include metadata section (default: True)",
                    "default": True
                }
            },
            "required": ["sheet_json"]
        },
        "outputSchema": {
            "type": "object",
            "properties": {
                "status": {
                    "type": "string",
                    "enum": ["success", "error"],
                    "description": "Operation status"
                },
                "html_content": {
                    "type": "string",
                    "description": "Generated HTML content (present when status is success)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Name of the converted sheet"
                },
                "metadata": {
                    "type": "object",
                    "description": "Conversion metadata"
                },
                "error_type": {
                    "type": "string",
                    "description": "Type of error (present when status is error)"
                },
                "error_message": {
                    "type": "string",
                    "description": "Error description (present when status is error)"
                }
            },
            "required": ["status"]
        }
    }
}


def get_tool_definition(tool_name: str) -> Dict[str, Any]:
    """
    Get the JSON Schema definition for a specific tool.
    
    Args:
        tool_name: Name of the tool
        
    Returns:
        Tool definition dictionary
        
    Raises:
        KeyError: If tool name is not found
    """
    if tool_name not in TOOL_DEFINITIONS:
        raise KeyError(f"Tool '{tool_name}' not found. Available tools: {list(TOOL_DEFINITIONS.keys())}")
    
    return TOOL_DEFINITIONS[tool_name]


def get_all_tool_definitions() -> Dict[str, Dict[str, Any]]:
    """
    Get all tool definitions.
    
    Returns:
        Dictionary of all tool definitions
    """
    return TOOL_DEFINITIONS.copy()


def validate_tool_input(tool_name: str, input_data: Dict[str, Any]) -> bool:
    """
    Validate input data against tool schema (basic validation).
    
    Args:
        tool_name: Name of the tool
        input_data: Input data to validate
        
    Returns:
        True if validation passes
        
    Raises:
        KeyError: If tool name is not found
        ValueError: If validation fails
    """
    tool_def = get_tool_definition(tool_name)
    input_schema = tool_def["inputSchema"]
    required_fields = input_schema.get("required", [])
    
    # Check required fields
    for field in required_fields:
        if field not in input_data:
            raise ValueError(f"Missing required field '{field}' for tool '{tool_name}'")
    
    return True
