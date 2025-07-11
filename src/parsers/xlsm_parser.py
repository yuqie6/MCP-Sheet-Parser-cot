import openpyxl
from typing import Optional

from src.models.table_model import Sheet, Row, Cell, Style
from src.parsers.base_parser import BaseParser


class XlsmParser(BaseParser):
    """Parses .xlsm files using openpyxl (same as xlsx but with macros)."""

    def parse(self, file_path: str, sheet_name: Optional[str] = None) -> Sheet:
        """Parses the given .xlsm file and returns a Sheet object."""
        # Load workbook with macro support
        workbook = openpyxl.load_workbook(file_path, keep_vba=True)
        
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found in the workbook")
            worksheet = workbook[sheet_name]
        else:
            worksheet = workbook.active

        sheet_data = Sheet(name=worksheet.title)
        
        for r_idx, row in enumerate(worksheet.iter_rows()):
            row_data = Row()
            for c_idx, cell in enumerate(row):
                # Enhanced style extraction for XLSM files
                font = cell.font
                fill = cell.fill
                alignment = cell.alignment
                border = cell.border
                
                style = Style(
                    # Basic formatting
                    bold=font.bold,
                    italic=font.italic,
                    underline=font.underline != 'none' if font.underline else False,
                    font_color=self._get_color_hex(font.color),
                    fill_color=self._get_color_hex(fill.fgColor) if fill.fgColor else None,
                    
                    # Extended formatting
                    font_size=font.size,
                    font_name=font.name,
                    alignment=alignment.horizontal,
                    vertical_alignment=alignment.vertical,
                    
                    # Border information (simplified)
                    border_style=self._get_border_style(border),
                )
                
                cell_data = Cell(
                    value=cell.value,
                    row=r_idx,
                    column=c_idx,
                    style=style
                )
                row_data.cells.append(cell_data)
            sheet_data.rows.append(row_data)
        
        return sheet_data
    
    def _get_color_hex(self, color) -> Optional[str]:
        """Convert openpyxl color to hex string."""
        if color and hasattr(color, 'rgb') and color.rgb:
            try:
                if color.rgb.startswith('FF'):
                    return f"#{color.rgb[2:]}"
                else:
                    return f"#{color.rgb}"
            except (AttributeError, TypeError):
                pass
        return None
    
    def _get_border_style(self, border) -> Optional[str]:
        """Extract border style information."""
        if not border:
            return None
        
        # Check if any border side has a style
        sides = [border.left, border.right, border.top, border.bottom]
        styles = []
        
        for side in sides:
            if side and side.style and side.style != 'none':
                styles.append(side.style)
        
        if styles:
            # Return the most common style or the first one
            return styles[0]
        
        return None
