import os
from typing import Dict, List, Any, Optional
from jinja2 import Environment, FileSystemLoader

from src.models.table_model import Sheet, Row, Cell
from src.converters.style_converter import StyleConverter


class CellData:
    """Helper class to represent cell data for template rendering."""
    
    def __init__(self, value: Any, style: str = "", colspan: int = 1, rowspan: int = 1, 
                 is_merged: bool = False, is_merged_continuation: bool = False, tag: str = "td"):
        self.value = value
        self.style = style
        self.colspan = colspan
        self.rowspan = rowspan
        self.is_merged = is_merged
        self.is_merged_continuation = is_merged_continuation
        self.tag = tag


class HtmlConverter:
    """Converts Sheet objects to HTML using Jinja2 templates."""
    
    def __init__(self):
        self.style_converter = StyleConverter()
        
        # Set up Jinja2 environment
        template_dir = os.path.join(os.path.dirname(__file__), 'templates')
        self.env = Environment(loader=FileSystemLoader(template_dir))
    
    def convert_sheet_to_html(self, sheet: Sheet, include_styles: bool = True, 
                            include_metadata: bool = True) -> str:
        """
        Convert a Sheet object to HTML.
        
        Args:
            sheet: Sheet object to convert
            include_styles: Whether to include styling information
            include_metadata: Whether to include metadata section
            
        Returns:
            HTML string representation of the sheet
        """
        # Process merged cells information
        merged_ranges = self._parse_merged_cells(sheet.merged_cells)
        
        # Convert rows to template-friendly format
        template_rows = self._convert_rows_for_template(
            sheet.rows, merged_ranges, include_styles
        )
        
        # Prepare metadata
        metadata = None
        if include_metadata:
            metadata = self._generate_metadata(sheet)
        
        # Render template
        template = self.env.get_template('table.html')
        html_content = template.render(
            sheet_name=sheet.name,
            rows=template_rows,
            metadata=metadata,
            default_css=self.style_converter.get_table_css()
        )
        
        return html_content
    
    def _parse_merged_cells(self, merged_cells: List[str]) -> List[Dict[str, Any]]:
        """
        Parse merged cell ranges from Excel-style notation.
        
        Args:
            merged_cells: List of merged cell ranges (e.g., ['A1:B2', 'C3:D4'])
            
        Returns:
            List of dictionaries with parsed range information
        """
        ranges = []
        for cell_range in merged_cells:
            if ':' in cell_range:
                start_cell, end_cell = cell_range.split(':')
                start_col, start_row = self._parse_cell_reference(start_cell)
                end_col, end_row = self._parse_cell_reference(end_cell)
                
                ranges.append({
                    'start_row': start_row,
                    'start_col': start_col,
                    'end_row': end_row,
                    'end_col': end_col,
                    'colspan': end_col - start_col + 1,
                    'rowspan': end_row - start_row + 1
                })
        
        return ranges
    
    def _parse_cell_reference(self, cell_ref: str) -> tuple:
        """
        Parse Excel-style cell reference (e.g., 'A1') to (column, row) indices.
        
        Args:
            cell_ref: Excel-style cell reference
            
        Returns:
            Tuple of (column_index, row_index) (0-based)
        """
        col_str = ""
        row_str = ""
        
        for char in cell_ref:
            if char.isalpha():
                col_str += char
            else:
                row_str += char
        
        # Convert column letters to index (A=0, B=1, etc.)
        col_index = 0
        for char in col_str:
            col_index = col_index * 26 + (ord(char.upper()) - ord('A') + 1)
        col_index -= 1  # Convert to 0-based
        
        # Convert row number to index (1-based to 0-based)
        row_index = int(row_str) - 1 if row_str else 0
        
        return col_index, row_index
    
    def _convert_rows_for_template(self, rows: List[Row], merged_ranges: List[Dict], 
                                 include_styles: bool) -> List[Dict[str, Any]]:
        """
        Convert Sheet rows to template-friendly format.
        
        Args:
            rows: List of Row objects
            merged_ranges: List of merged cell range information
            include_styles: Whether to include styling
            
        Returns:
            List of dictionaries suitable for template rendering
        """
        template_rows = []
        
        for row_idx, row in enumerate(rows):
            template_cells = []
            
            for col_idx, cell in enumerate(row.cells):
                # Check if this cell is part of a merged range
                merge_info = self._get_merge_info(row_idx, col_idx, merged_ranges)
                
                # Skip cells that are continuation of merged cells
                if merge_info and merge_info['is_continuation']:
                    continue
                
                # Convert style to CSS
                style_css = ""
                if include_styles and cell.style:
                    style_css = self.style_converter.style_to_css(cell.style)
                
                # Determine cell tag (th for header rows, td for data)
                cell_tag = "th" if row_idx == 0 else "td"
                
                # Create cell data
                cell_data = CellData(
                    value=cell.value,
                    style=style_css,
                    colspan=merge_info['colspan'] if merge_info else 1,
                    rowspan=merge_info['rowspan'] if merge_info else 1,
                    is_merged=bool(merge_info),
                    is_merged_continuation=False,
                    tag=cell_tag
                )
                
                template_cells.append(cell_data)
            
            template_rows.append({'cells': template_cells})
        
        return template_rows
    
    def _get_merge_info(self, row_idx: int, col_idx: int, 
                       merged_ranges: List[Dict]) -> Optional[Dict]:
        """
        Get merge information for a specific cell.
        
        Args:
            row_idx: Row index of the cell
            col_idx: Column index of the cell
            merged_ranges: List of merged cell ranges
            
        Returns:
            Dictionary with merge information or None
        """
        for merge_range in merged_ranges:
            if (merge_range['start_row'] <= row_idx <= merge_range['end_row'] and
                merge_range['start_col'] <= col_idx <= merge_range['end_col']):
                
                # Check if this is the top-left cell (main cell) or continuation
                is_continuation = not (row_idx == merge_range['start_row'] and 
                                     col_idx == merge_range['start_col'])
                
                return {
                    'colspan': merge_range['colspan'],
                    'rowspan': merge_range['rowspan'],
                    'is_continuation': is_continuation
                }
        
        return None
    
    def _generate_metadata(self, sheet: Sheet) -> Dict[str, Any]:
        """
        Generate metadata information for the sheet.
        
        Args:
            sheet: Sheet object
            
        Returns:
            Dictionary with metadata information
        """
        total_rows = len(sheet.rows)
        total_columns = max(len(row.cells) for row in sheet.rows) if sheet.rows else 0
        merged_cells_count = len(sheet.merged_cells)
        
        return {
            'total_rows': total_rows,
            'total_columns': total_columns,
            'merged_cells_count': merged_cells_count,
            'file_format': 'Unknown'  # This could be enhanced to track source format
        }
