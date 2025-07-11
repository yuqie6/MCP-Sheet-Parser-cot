"""
JSON Converter Module

This module provides functionality to convert Sheet objects to JSON format,
creating LLM-friendly data structures that are compact and comprehensive.
"""

import json
from dataclasses import asdict
from typing import Dict, Any, List, Optional

from src.models.table_model import Sheet, Row, Cell, Style


class JSONConverter:
    """
    Converts Sheet objects to JSON format for LLM consumption.
    
    The JSON format is designed to be:
    - Compact: Minimal redundancy and efficient structure
    - Comprehensive: Contains all data and style information
    - LLM-friendly: Easy to parse and analyze by language models
    """
    
    def convert(self, sheet: Sheet) -> Dict[str, Any]:
        """
        Convert a Sheet object to JSON format.
        
        Args:
            sheet: The Sheet object to convert
            
        Returns:
            A dictionary representing the sheet in JSON format
        """
        if not sheet:
            raise ValueError("Sheet cannot be None")
        
        # Calculate dimensions
        row_count = len(sheet.rows)
        col_count = len(sheet.rows[0].cells) if sheet.rows else 0
        
        # Build JSON structure
        json_data = {
            'metadata': {
                'name': sheet.name or 'Untitled',
                'rows': row_count,
                'cols': col_count,
                'has_merged_cells': bool(sheet.merged_cells),
                'merged_cells_count': len(sheet.merged_cells) if sheet.merged_cells else 0
            },
            'data': self._convert_rows_to_json(sheet.rows),
            'merged_cells': sheet.merged_cells or [],
            'styles': self._extract_unique_styles(sheet.rows)
        }
        
        return json_data
    
    def _convert_rows_to_json(self, rows: List[Row]) -> List[Dict[str, Any]]:
        """
        Convert rows to JSON format.
        
        Args:
            rows: List of Row objects
            
        Returns:
            List of row dictionaries
        """
        json_rows = []
        
        for row_idx, row in enumerate(rows):
            row_data = {
                'row': row_idx,
                'cells': []
            }
            
            for cell_idx, cell in enumerate(row.cells):
                cell_data = {
                    'col': cell_idx,
                    'value': self._serialize_cell_value(cell.value),
                    'style_id': self._get_style_id(cell.style) if cell.style else None
                }
                
                # Add colspan/rowspan if not default
                if hasattr(cell, 'col_span') and cell.col_span > 1:
                    cell_data['col_span'] = cell.col_span
                if hasattr(cell, 'row_span') and cell.row_span > 1:
                    cell_data['row_span'] = cell.row_span
                
                row_data['cells'].append(cell_data)
            
            json_rows.append(row_data)
        
        return json_rows
    
    def _serialize_cell_value(self, value: Any) -> Any:
        """
        Serialize cell value to JSON-compatible format.
        
        Args:
            value: The cell value to serialize
            
        Returns:
            JSON-compatible value
        """
        if value is None:
            return None
        
        # Handle different data types
        if isinstance(value, (str, int, float, bool)):
            return value
        
        # Convert other types to string
        return str(value)
    
    def _extract_unique_styles(self, rows: List[Row]) -> Dict[str, Dict[str, Any]]:
        """
        Extract unique styles and create a style dictionary.
        
        This reduces JSON size by avoiding style duplication.
        
        Args:
            rows: List of Row objects
            
        Returns:
            Dictionary mapping style IDs to style objects
        """
        unique_styles = {}
        style_counter = 1
        
        for row in rows:
            for cell in row.cells:
                if cell.style:
                    style_id = self._get_style_id(cell.style)
                    if style_id not in unique_styles:
                        unique_styles[style_id] = self._style_to_dict(cell.style)
        
        return unique_styles
    
    def _get_style_id(self, style: Style) -> str:
        """
        Generate a unique ID for a style based on its properties.
        
        Args:
            style: The Style object
            
        Returns:
            Unique style identifier
        """
        if not style:
            return "default"
        
        # Create a hash-like ID based on style properties
        style_signature = (
            f"{style.bold}_{style.italic}_{style.underline}_"
            f"{style.font_color}_{style.background_color}_"
            f"{style.text_align}_{style.vertical_align}_"
            f"{style.font_size}_{style.font_name}_"
            f"{style.border_top}_{style.border_bottom}_"
            f"{style.border_left}_{style.border_right}_"
            f"{style.wrap_text}_{style.number_format}"
        )
        
        # Generate a shorter ID
        return f"style_{abs(hash(style_signature)) % 10000}"
    
    def _style_to_dict(self, style: Style) -> Dict[str, Any]:
        """
        Convert Style object to dictionary.
        
        Args:
            style: The Style object to convert
            
        Returns:
            Dictionary representation of the style
        """
        if not style:
            return {}
        
        style_dict = asdict(style)
        
        # Remove None values and empty strings to keep JSON compact
        return {k: v for k, v in style_dict.items() if v is not None and v != ""}
    
    def to_json_string(self, sheet: Sheet, indent: Optional[int] = None) -> str:
        """
        Convert Sheet to JSON string.
        
        Args:
            sheet: The Sheet object to convert
            indent: JSON indentation (None for compact format)
            
        Returns:
            JSON string representation
        """
        json_data = self.convert(sheet)
        return json.dumps(json_data, indent=indent, ensure_ascii=False)
    
    def estimate_json_size(self, sheet: Sheet) -> Dict[str, int]:
        """
        Estimate the size of the JSON output.
        
        Args:
            sheet: The Sheet object to analyze
            
        Returns:
            Dictionary with size estimates
        """
        json_data = self.convert(sheet)
        json_string = json.dumps(json_data, ensure_ascii=False)
        
        return {
            'total_characters': len(json_string),
            'total_bytes': len(json_string.encode('utf-8')),
            'rows': len(json_data['data']),
            'unique_styles': len(json_data['styles']),
            'cells': sum(len(row['cells']) for row in json_data['data'])
        }
