import os
from typing import Dict, List, Any, Optional, Union
import logging

from src.parsers.parser_factory import get_parser
from src.converters.html_converter import HtmlConverter
from src.models.table_model import Sheet, Workbook, Style


class ParsingServiceError(Exception):
    """Custom exception for parsing service errors."""
    pass


class ParsingService:
    """
    Business service layer that orchestrates the complete workflow
    from file parsing to HTML generation.
    """
    
    def __init__(self):
        self.html_converter = HtmlConverter()
        self.logger = logging.getLogger(__name__)
    
    def convert_to_html(self, file_path: str, sheet_name: Optional[str] = None, 
                       include_styles: bool = True, include_metadata: bool = True) -> Dict[str, Any]:
        """
        Convert a spreadsheet file to HTML.
        
        Args:
            file_path: Path to the spreadsheet file
            sheet_name: Name of the sheet to convert (optional)
            include_styles: Whether to include styling information
            include_metadata: Whether to include metadata section
            
        Returns:
            Dictionary containing HTML content and metadata
            
        Raises:
            ParsingServiceError: If parsing or conversion fails
        """
        try:
            # Validate input parameters
            self._validate_file_path(file_path)
            
            # Get appropriate parser
            parser = get_parser(file_path)
            self.logger.info(f"Using parser: {type(parser).__name__} for file: {file_path}")
            
            # Parse the file
            sheet = parser.parse(file_path, sheet_name)
            self.logger.info(f"Successfully parsed sheet: {sheet.name}")
            
            # Convert to HTML
            html_content = self.html_converter.convert_sheet_to_html(
                sheet, include_styles, include_metadata
            )
            
            # Prepare response
            result = {
                "status": "success",
                "html_content": html_content,
                "sheet_name": sheet.name,
                "metadata": {
                    "file_path": file_path,
                    "file_format": self._get_file_format(file_path),
                    "total_rows": len(sheet.rows),
                    "total_columns": max(len(row.cells) for row in sheet.rows) if sheet.rows else 0,
                    "merged_cells_count": len(sheet.merged_cells),
                    "include_styles": include_styles,
                    "include_metadata": include_metadata
                }
            }
            
            self.logger.info(f"Successfully converted to HTML. Length: {len(html_content)}")
            return result
            
        except Exception as e:
            error_msg = f"Failed to convert file {file_path} to HTML: {str(e)}"
            self.logger.error(error_msg)
            raise ParsingServiceError(error_msg) from e
    
    def get_sheet_metadata(self, file_path: str) -> Dict[str, Any]:
        """
        Get metadata about a spreadsheet file without full parsing.
        
        Args:
            file_path: Path to the spreadsheet file
            
        Returns:
            Dictionary containing file and sheet metadata
            
        Raises:
            ParsingServiceError: If metadata extraction fails
        """
        try:
            # Validate input
            self._validate_file_path(file_path)
            
            # Get file-level metadata
            file_stats = os.stat(file_path)
            file_format = self._get_file_format(file_path)
            
            # Try to get sheet names (this might require partial parsing)
            sheet_names = self._get_sheet_names(file_path)
            
            result = {
                "status": "success",
                "file_path": file_path,
                "file_format": file_format,
                "file_size": file_stats.st_size,
                "last_modified": file_stats.st_mtime,
                "sheet_names": sheet_names,
                "total_sheets": len(sheet_names)
            }
            
            self.logger.info(f"Successfully extracted metadata for: {file_path}")
            return result
            
        except Exception as e:
            error_msg = f"Failed to get metadata for file {file_path}: {str(e)}"
            self.logger.error(error_msg)
            raise ParsingServiceError(error_msg) from e
    
    def parse_sheet_to_json(self, file_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Parse a sheet and return structured JSON data.
        
        Args:
            file_path: Path to the spreadsheet file
            sheet_name: Name of the sheet to parse (optional)
            
        Returns:
            Dictionary containing structured sheet data
            
        Raises:
            ParsingServiceError: If parsing fails
        """
        try:
            # Validate input
            self._validate_file_path(file_path)
            
            # Get parser and parse
            parser = get_parser(file_path)
            sheet = parser.parse(file_path, sheet_name)
            
            # Convert to JSON-serializable format
            json_data = self._sheet_to_json(sheet)
            
            result = {
                "status": "success",
                "sheet_data": json_data,
                "metadata": {
                    "file_path": file_path,
                    "file_format": self._get_file_format(file_path),
                    "sheet_name": sheet.name,
                    "total_rows": len(sheet.rows),
                    "total_columns": max(len(row.cells) for row in sheet.rows) if sheet.rows else 0,
                    "merged_cells_count": len(sheet.merged_cells)
                }
            }
            
            self.logger.info(f"Successfully parsed sheet to JSON: {sheet.name}")
            return result
            
        except Exception as e:
            error_msg = f"Failed to parse sheet to JSON: {str(e)}"
            self.logger.error(error_msg)
            raise ParsingServiceError(error_msg) from e
    
    def convert_json_to_html(self, sheet_json: Dict[str, Any], 
                           include_styles: bool = True, include_metadata: bool = True) -> Dict[str, Any]:
        """
        Convert JSON sheet data to HTML.
        
        Args:
            sheet_json: JSON representation of sheet data
            include_styles: Whether to include styling information
            include_metadata: Whether to include metadata section
            
        Returns:
            Dictionary containing HTML content
            
        Raises:
            ParsingServiceError: If conversion fails
        """
        try:
            # Convert JSON back to Sheet object
            sheet = self._json_to_sheet(sheet_json)
            
            # Convert to HTML
            html_content = self.html_converter.convert_sheet_to_html(
                sheet, include_styles, include_metadata
            )
            
            result = {
                "status": "success",
                "html_content": html_content,
                "sheet_name": sheet.name,
                "metadata": {
                    "total_rows": len(sheet.rows),
                    "total_columns": max(len(row.cells) for row in sheet.rows) if sheet.rows else 0,
                    "merged_cells_count": len(sheet.merged_cells),
                    "include_styles": include_styles,
                    "include_metadata": include_metadata
                }
            }
            
            self.logger.info(f"Successfully converted JSON to HTML")
            return result
            
        except Exception as e:
            error_msg = f"Failed to convert JSON to HTML: {str(e)}"
            self.logger.error(error_msg)
            raise ParsingServiceError(error_msg) from e
    
    def _validate_file_path(self, file_path: str) -> None:
        """Validate file path and accessibility."""
        if not file_path:
            raise ParsingServiceError("File path cannot be empty")
        
        if not os.path.exists(file_path):
            raise ParsingServiceError(f"File does not exist: {file_path}")
        
        if not os.path.isfile(file_path):
            raise ParsingServiceError(f"Path is not a file: {file_path}")
        
        if not os.access(file_path, os.R_OK):
            raise ParsingServiceError(f"File is not readable: {file_path}")
    
    def _get_file_format(self, file_path: str) -> str:
        """Get file format from extension."""
        _, extension = os.path.splitext(file_path)
        return extension.lower().lstrip('.')
    
    def _get_sheet_names(self, file_path: str) -> List[str]:
        """Get sheet names from file (may require partial parsing)."""
        try:
            # For now, we'll do a quick parse to get sheet names
            # This could be optimized for specific formats in the future
            parser = get_parser(file_path)

            # For most formats, we need to do a quick parse to get sheet names
            # This is a simplified approach - could be enhanced per format
            try:
                # Try to parse the first sheet to get at least one name
                sheet = parser.parse(file_path)
                return [sheet.name] if sheet.name else ["Sheet1"]
            except Exception:
                # If parsing fails, return a default
                return ["Sheet1"]

        except Exception:
            # If we can't get sheet names, return a default
            return ["Sheet1"]

    def _sheet_to_json(self, sheet: Sheet) -> Dict[str, Any]:
        """Convert Sheet object to JSON-serializable format."""
        rows_data = []

        for row in sheet.rows:
            cells_data = []
            for cell in row.cells:
                cell_data = {
                    "value": cell.value,
                    "row": cell.row,
                    "column": cell.column,
                    "style": self._style_to_dict(cell.style) if cell.style else None
                }
                cells_data.append(cell_data)
            rows_data.append({"cells": cells_data})

        return {
            "name": sheet.name,
            "rows": rows_data,
            "merged_cells": sheet.merged_cells
        }

    def _json_to_sheet(self, sheet_json: Dict[str, Any]) -> Sheet:
        """Convert JSON data back to Sheet object."""
        from src.models.table_model import Row, Cell

        sheet = Sheet(name=sheet_json.get("name", "Sheet1"))
        sheet.merged_cells = sheet_json.get("merged_cells", [])

        for row_data in sheet_json.get("rows", []):
            row = Row()
            for cell_data in row_data.get("cells", []):
                style = None
                if cell_data.get("style"):
                    style = self._dict_to_style(cell_data["style"])

                cell = Cell(
                    value=cell_data.get("value"),
                    row=cell_data.get("row", 0),
                    column=cell_data.get("column", 0),
                    style=style
                )
                row.cells.append(cell)
            sheet.rows.append(row)

        return sheet

    def _style_to_dict(self, style: Style) -> Dict[str, Any]:
        """Convert Style object to dictionary."""
        return {
            "bold": style.bold,
            "italic": style.italic,
            "underline": style.underline,
            "font_color": style.font_color,
            "fill_color": style.fill_color,
            "font_size": style.font_size,
            "font_name": style.font_name,
            "alignment": style.alignment,
            "vertical_alignment": style.vertical_alignment,
            "border_style": style.border_style,
            "border_color": style.border_color,
            "border_width": style.border_width,
            "text_decoration": style.text_decoration,
            "width": style.width,
            "height": style.height,
            "padding": style.padding
        }

    def _dict_to_style(self, style_dict: Dict[str, Any]) -> Style:
        """Convert dictionary back to Style object."""

        return Style(
            bold=style_dict.get("bold"),
            italic=style_dict.get("italic"),
            underline=style_dict.get("underline"),
            font_color=style_dict.get("font_color"),
            fill_color=style_dict.get("fill_color"),
            font_size=style_dict.get("font_size"),
            font_name=style_dict.get("font_name"),
            alignment=style_dict.get("alignment"),
            vertical_alignment=style_dict.get("vertical_alignment"),
            border_style=style_dict.get("border_style"),
            border_color=style_dict.get("border_color"),
            border_width=style_dict.get("border_width"),
            text_decoration=style_dict.get("text_decoration"),
            width=style_dict.get("width"),
            height=style_dict.get("height"),
            padding=style_dict.get("padding")
        )
