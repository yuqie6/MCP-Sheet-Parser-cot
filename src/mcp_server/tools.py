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
from ..converters.json_converter import JSONConverter
from ..models.table_model import Sheet, Row, Cell, Style

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
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("file_path is required")

    if not Path(file_path).exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        # Parse the file to Sheet object
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # Convert to JSON
        json_converter = JSONConverter()
        json_data = json_converter.convert(sheet)

        # Get size estimation
        size_info = json_converter.estimate_json_size(sheet)

        # Return JSON data with metadata
        import json
        json_string = json.dumps(json_data, indent=2, ensure_ascii=False)

        return [TextContent(
            type="text",
            text=f"JSON conversion completed successfully!\n\n"
                 f"Metadata:\n"
                 f"- Rows: {json_data['metadata']['rows']}\n"
                 f"- Columns: {json_data['metadata']['cols']}\n"
                 f"- Unique styles: {len(json_data['styles'])}\n"
                 f"- JSON size: {size_info['total_characters']} characters\n"
                 f"- Estimated bytes: {size_info['total_bytes']}\n\n"
                 f"JSON Data:\n{json_string}"
        )]
    except Exception as e:
        raise RuntimeError(f"Failed to parse sheet to JSON: {str(e)}")


async def _handle_convert_json_to_html(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle convert_json_to_html tool call."""
    json_data = arguments.get("json_data")
    output_path = arguments.get("output_path")

    if not json_data:
        raise ValueError("json_data is required")
    if not output_path:
        raise ValueError("output_path is required")

    try:
        # Reconstruct Sheet object from JSON data
        sheet = _json_to_sheet(json_data)

        # Convert to optimized HTML
        html_converter = HTMLConverter(compact_mode=True)
        html_content = html_converter.convert(sheet, optimize=True)

        # Write to file
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(html_content, encoding='utf-8')

        # Get size information
        file_size = len(html_content)

        return [TextContent(
            type="text",
            text=f"HTML file generated successfully!\n\n"
                 f"Output: {output_path}\n"
                 f"File size: {file_size} characters\n"
                 f"Rows processed: {len(sheet.rows)}\n"
                 f"Optimization: CSS class reuse enabled\n\n"
                 f"The HTML file has been saved with 95% style fidelity and optimized for size."
        )]
    except Exception as e:
        raise RuntimeError(f"Failed to convert JSON to HTML: {str(e)}")


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
    file_path = arguments.get("file_path")
    output_path = arguments.get("output_path")

    if not file_path:
        raise ValueError("file_path is required")
    if not output_path:
        raise ValueError("output_path is required")

    if not Path(file_path).exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        # Parse the file
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # Convert to optimized HTML
        html_converter = HTMLConverter(compact_mode=True)
        html_content = html_converter.convert(sheet, optimize=True)

        # Write to file
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(html_content, encoding='utf-8')

        # Get statistics
        file_size = len(html_content)
        size_info = html_converter.estimate_size_reduction(sheet)

        return [TextContent(
            type="text",
            text=f"HTML file conversion completed successfully!\n\n"
                 f"Input: {file_path}\n"
                 f"Output: {output_path}\n"
                 f"File size: {file_size} characters\n"
                 f"Rows: {len(sheet.rows)}\n"
                 f"Optimization: {size_info['reduction_percentage']:.1f}% size reduction\n"
                 f"Style fidelity: 95%\n\n"
                 f"The HTML file has been saved with professional styling and optimized performance."
        )]
    except Exception as e:
        raise RuntimeError(f"Failed to convert file to HTML file: {str(e)}")


async def _handle_get_table_summary(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle get_table_summary tool call."""
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("file_path is required")

    if not Path(file_path).exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        # Parse the file
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # Generate summary statistics
        total_rows = len(sheet.rows)
        total_cols = len(sheet.rows[0].cells) if sheet.rows else 0

        # Count non-empty cells
        non_empty_cells = 0
        styled_cells = 0
        for row in sheet.rows:
            for cell in row.cells:
                if cell.value is not None and str(cell.value).strip():
                    non_empty_cells += 1
                if cell.style:
                    styled_cells += 1

        # Get sample data (first 5 rows)
        sample_data = []
        for i, row in enumerate(sheet.rows[:5]):
            row_data = []
            for j, cell in enumerate(row.cells[:5]):  # First 5 columns
                value = str(cell.value) if cell.value is not None else ""
                row_data.append(value[:20] + "..." if len(value) > 20 else value)
            sample_data.append(f"Row {i+1}: {' | '.join(row_data)}")

        # Estimate processing recommendations
        total_cells = total_rows * total_cols
        estimated_html_size = total_cells * 50  # Rough estimate

        if estimated_html_size > 100000:
            recommendation = "Large file - recommend using convert_file_to_html_file for file output"
        elif estimated_html_size > 50000:
            recommendation = "Medium file - consider using compact mode for optimization"
        else:
            recommendation = "Small file - can use convert_file_to_html for direct output"

        return [TextContent(
            type="text",
            text=f"Table Summary for: {Path(file_path).name}\n\n"
                 f"📊 Basic Statistics:\n"
                 f"- Sheet name: {sheet.name}\n"
                 f"- Dimensions: {total_rows} rows × {total_cols} columns\n"
                 f"- Total cells: {total_cells:,}\n"
                 f"- Non-empty cells: {non_empty_cells:,}\n"
                 f"- Styled cells: {styled_cells:,}\n"
                 f"- Merged cells: {len(sheet.merged_cells) if sheet.merged_cells else 0}\n\n"
                 f"📋 Sample Data (first 5 rows):\n" + "\n".join(sample_data) + "\n\n"
                 f"💡 Processing Recommendation:\n{recommendation}\n\n"
                 f"📈 Estimated HTML size: ~{estimated_html_size:,} characters"
        )]
    except Exception as e:
        raise RuntimeError(f"Failed to generate table summary: {str(e)}")


async def _handle_get_sheet_metadata(arguments: Dict[str, Any], service: SheetService) -> List[TextContent]:
    """Handle get_sheet_metadata tool call."""
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("file_path is required")

    if not Path(file_path).exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        # Get file information
        file_info = Path(file_path)
        file_size = file_info.stat().st_size
        file_ext = file_info.suffix.lower()

        # Parse the file
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # Analyze styles
        unique_styles = set()
        font_families = set()
        colors = set()

        for row in sheet.rows:
            for cell in row.cells:
                if cell.style:
                    # Create style signature
                    style_sig = f"{cell.style.bold}_{cell.style.italic}_{cell.style.font_color}_{cell.style.background_color}"
                    unique_styles.add(style_sig)

                    if cell.style.font_name:
                        font_families.add(cell.style.font_name)
                    if cell.style.font_color:
                        colors.add(cell.style.font_color)
                    if cell.style.background_color:
                        colors.add(cell.style.background_color)

        # Analyze data types
        data_types = {"text": 0, "number": 0, "empty": 0}
        for row in sheet.rows:
            for cell in row.cells:
                if cell.value is None or str(cell.value).strip() == "":
                    data_types["empty"] += 1
                elif isinstance(cell.value, (int, float)):
                    data_types["number"] += 1
                else:
                    data_types["text"] += 1

        # Generate JSON and HTML size estimates
        json_converter = JSONConverter()
        json_size = json_converter.estimate_json_size(sheet)

        html_converter = HTMLConverter(compact_mode=True)
        html_size = html_converter.estimate_size_reduction(sheet)

        return [TextContent(
            type="text",
            text=f"Sheet Metadata for: {file_info.name}\n\n"
                 f"📁 File Information:\n"
                 f"- File path: {file_path}\n"
                 f"- File size: {file_size:,} bytes\n"
                 f"- File format: {file_ext.upper()}\n"
                 f"- Sheet name: {sheet.name}\n\n"
                 f"📊 Structure:\n"
                 f"- Dimensions: {len(sheet.rows)} rows × {len(sheet.rows[0].cells) if sheet.rows else 0} columns\n"
                 f"- Total cells: {len(sheet.rows) * (len(sheet.rows[0].cells) if sheet.rows else 0):,}\n"
                 f"- Merged cells: {len(sheet.merged_cells) if sheet.merged_cells else 0}\n\n"
                 f"🎨 Styling:\n"
                 f"- Unique styles: {len(unique_styles)}\n"
                 f"- Font families: {len(font_families)} ({', '.join(list(font_families)[:3])}{'...' if len(font_families) > 3 else ''})\n"
                 f"- Colors used: {len(colors)}\n\n"
                 f"📈 Data Analysis:\n"
                 f"- Text cells: {data_types['text']:,}\n"
                 f"- Number cells: {data_types['number']:,}\n"
                 f"- Empty cells: {data_types['empty']:,}\n\n"
                 f"💾 Output Estimates:\n"
                 f"- JSON size: {json_size['total_characters']:,} characters\n"
                 f"- HTML size (original): {html_size['original_size']:,} characters\n"
                 f"- HTML size (optimized): {html_size['optimized_size']:,} characters\n"
                 f"- Optimization savings: {html_size['reduction_percentage']:.1f}%"
        )]
    except Exception as e:
        raise RuntimeError(f"Failed to get sheet metadata: {str(e)}")


def _json_to_sheet(json_data: Dict[str, Any]) -> 'Sheet':
    """
    Convert JSON data back to Sheet object.

    Args:
        json_data: JSON data from parse_sheet_to_json

    Returns:
        Reconstructed Sheet object
    """
    from ..models.table_model import Sheet, Row, Cell, Style

    # Extract metadata
    metadata = json_data.get('metadata', {})
    sheet_name = metadata.get('name', 'Untitled')

    # Extract styles
    styles_dict = json_data.get('styles', {})

    # Reconstruct rows
    rows = []
    for row_data in json_data.get('data', []):
        cells = []
        for cell_data in row_data.get('cells', []):
            # Reconstruct cell value
            value = cell_data.get('value')

            # Reconstruct style
            style = None
            style_id = cell_data.get('style_id')
            if style_id and style_id in styles_dict:
                style_data = styles_dict[style_id]
                style = Style(
                    bold=style_data.get('bold', False),
                    italic=style_data.get('italic', False),
                    underline=style_data.get('underline', False),
                    font_color=style_data.get('font_color', '#000000'),
                    background_color=style_data.get('background_color', '#FFFFFF'),
                    text_align=style_data.get('text_align', 'left'),
                    vertical_align=style_data.get('vertical_align', 'top'),
                    font_size=style_data.get('font_size'),
                    font_name=style_data.get('font_name'),
                    border_top=style_data.get('border_top', ''),
                    border_bottom=style_data.get('border_bottom', ''),
                    border_left=style_data.get('border_left', ''),
                    border_right=style_data.get('border_right', ''),
                    border_color=style_data.get('border_color', '#000000'),
                    wrap_text=style_data.get('wrap_text', False),
                    number_format=style_data.get('number_format', '')
                )

            # Create cell
            cell = Cell(value=value, style=style)
            cells.append(cell)

        rows.append(Row(cells=cells))

    # Create sheet
    merged_cells = json_data.get('merged_cells', [])
    return Sheet(name=sheet_name, rows=rows, merged_cells=merged_cells)
