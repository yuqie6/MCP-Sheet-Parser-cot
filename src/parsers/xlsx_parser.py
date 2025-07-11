import openpyxl
from typing import Optional

from src.models.table_model import Sheet, Row, Cell, Style
from src.parsers.base_parser import BaseParser

class XlsxParser(BaseParser):
    """Parses .xlsx files using openpyxl."""

    def parse(self, file_path: str, sheet_name: Optional[str] = None) -> Sheet:
        """Parses the given .xlsx file and returns a Sheet object."""
        workbook = openpyxl.load_workbook(file_path)
        if sheet_name:
            worksheet = workbook[sheet_name]
        else:
            worksheet = workbook.active

        sheet_data = Sheet(name=worksheet.title)
        for r_idx, row in enumerate(worksheet.iter_rows()):
            row_data = Row()
            for c_idx, cell in enumerate(row):
                # Enhanced style extraction with all supported attributes
                font = cell.font
                fill = cell.fill
                alignment = cell.alignment
                border = cell.border

                style = Style(
                    # Basic text formatting (existing attributes)
                    bold=font.bold,
                    italic=font.italic,
                    underline=font.underline != 'none' if font.underline else False,
                    font_color=self._get_color_hex(font.color),
                    fill_color=self._get_color_hex(fill.fgColor) if fill.fgColor else None,

                    # Extended font properties
                    font_size=font.size,
                    font_name=font.name,

                    # Text alignment
                    alignment=alignment.horizontal,
                    vertical_alignment=alignment.vertical,

                    # Border styling
                    border_style=self._get_border_style(border),
                    border_color=self._get_border_color(border),
                    border_width=self._get_border_width(border),

                    # Additional text decoration
                    text_decoration=self._get_text_decoration(font),

                    # Cell dimensions (these would typically come from column/row settings)
                    width=None,  # Could be enhanced to get column width
                    height=None,  # Could be enhanced to get row height
                    padding=None  # Excel doesn't have direct padding, could use indent
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
        if not color:
            return None

        try:
            if hasattr(color, 'rgb') and color.rgb:
                # Handle RGB color
                rgb = color.rgb
                if rgb.startswith('FF'):
                    # Remove alpha channel if present
                    return f"#{rgb[2:]}"
                else:
                    return f"#{rgb}"
            elif hasattr(color, 'indexed') and color.indexed is not None:
                # Handle indexed colors (basic Excel colors)
                # This is a simplified mapping - could be expanded
                indexed_colors = {
                    0: "#000000",  # Black
                    1: "#FFFFFF",  # White
                    2: "#FF0000",  # Red
                    3: "#00FF00",  # Green
                    4: "#0000FF",  # Blue
                    5: "#FFFF00",  # Yellow
                    6: "#FF00FF",  # Magenta
                    7: "#00FFFF",  # Cyan
                }
                return indexed_colors.get(color.indexed)
        except (AttributeError, TypeError, ValueError):
            pass

        return None

    def _get_border_style(self, border) -> Optional[str]:
        """Extract border style information."""
        if not border:
            return None

        # Check all border sides and return the most prominent style
        sides = [border.left, border.right, border.top, border.bottom]
        styles = []

        for side in sides:
            if side and side.style and side.style != 'none':
                styles.append(side.style)

        if styles:
            # Return the first non-none style found
            # Could be enhanced to handle different styles per side
            return styles[0]

        return None

    def _get_border_color(self, border) -> Optional[str]:
        """Extract border color information."""
        if not border:
            return None

        # Check all border sides for color
        sides = [border.left, border.right, border.top, border.bottom]

        for side in sides:
            if side and side.color:
                color_hex = self._get_color_hex(side.color)
                if color_hex:
                    return color_hex

        return None

    def _get_border_width(self, border) -> Optional[str]:
        """Extract border width information."""
        if not border:
            return None

        # Check if any border side has a style (indicating presence of border)
        sides = [border.left, border.right, border.top, border.bottom]

        for side in sides:
            if side and side.style and side.style != 'none':
                # Map Excel border styles to CSS widths
                style_width_map = {
                    'thin': '1px',
                    'medium': '2px',
                    'thick': '3px',
                    'hair': '0.5px',
                    'dotted': '1px',
                    'dashed': '1px',
                    'double': '3px'
                }
                return style_width_map.get(side.style, '1px')

        return None

    def _get_text_decoration(self, font) -> Optional[str]:
        """Extract text decoration information."""
        decorations = []

        if font.underline and font.underline != 'none':
            decorations.append('underline')

        if hasattr(font, 'strike') and font.strike:
            decorations.append('line-through')

        if decorations:
            return ' '.join(decorations)

        return None
