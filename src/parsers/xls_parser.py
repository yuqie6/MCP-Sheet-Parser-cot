import xlrd
from typing import Optional

from src.models.table_model import Sheet, Row, Cell, Style
from src.parsers.base_parser import BaseParser


class XlsParser(BaseParser):
    """Parses .xls files using xlrd."""

    def parse(self, file_path: str, sheet_name: Optional[str] = None) -> Sheet:
        """Parses the given .xls file and returns a Sheet object."""
        workbook = xlrd.open_workbook(file_path, formatting_info=True)
        
        # Select the worksheet
        if sheet_name:
            try:
                worksheet = workbook.sheet_by_name(sheet_name)
            except xlrd.XLRDError:
                raise ValueError(f"Sheet '{sheet_name}' not found in the workbook")
        else:
            worksheet = workbook.sheet_by_index(0)

        sheet_data = Sheet(name=worksheet.name)
        
        # Get formatting information
        book_format_map = workbook.format_map
        book_font_map = workbook.font_map
        book_xf_list = workbook.xf_list
        
        for r_idx in range(worksheet.nrows):
            row_data = Row()
            for c_idx in range(worksheet.ncols):
                cell_value = worksheet.cell_value(r_idx, c_idx)
                cell_type = worksheet.cell_type(r_idx, c_idx)
                
                # Extract style information
                style = None
                try:
                    cell_xf_index = worksheet.cell_xf_index(r_idx, c_idx)
                    if cell_xf_index < len(book_xf_list):
                        xf = book_xf_list[cell_xf_index]
                        font = book_font_map.get(xf.font_index)
                        
                        if font:
                            style = Style(
                                bold=font.bold,
                                italic=font.italic,
                                underline=font.underline_type != 0,
                                font_size=font.height // 20,  # Convert from twips to points
                                font_name=font.name,
                                font_color=self._get_color_from_index(font.colour_index, workbook)
                            )
                except (IndexError, AttributeError):
                    # If style extraction fails, create a basic style
                    style = Style()
                
                # Convert cell value based on type
                if cell_type == xlrd.XL_CELL_EMPTY:
                    cell_value = None
                elif cell_type == xlrd.XL_CELL_DATE:
                    # Convert Excel date to Python date
                    try:
                        cell_value = xlrd.xldate_as_datetime(cell_value, workbook.datemode)
                    except xlrd.XLDateError:
                        pass  # Keep original value if conversion fails
                
                cell_data = Cell(
                    value=cell_value,
                    row=r_idx,
                    column=c_idx,
                    style=style
                )
                row_data.cells.append(cell_data)
            sheet_data.rows.append(row_data)
        
        return sheet_data
    
    def _get_color_from_index(self, color_index: int, workbook) -> Optional[str]:
        """Convert Excel color index to hex color string."""
        try:
            if hasattr(workbook, 'colour_map') and color_index in workbook.colour_map:
                rgb = workbook.colour_map[color_index]
                if rgb and len(rgb) >= 3:
                    return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        except (AttributeError, IndexError, TypeError):
            pass
        return None
