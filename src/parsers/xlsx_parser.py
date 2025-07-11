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
                # Basic style extraction (can be expanded)
                font = cell.font
                fill = cell.fill
                style = Style(
                    bold=font.bold,
                    italic=font.italic,
                    underline=font.underline,
                    font_color=f'#{font.color.rgb}' if font.color else None,
                    fill_color=f'#{fill.fgColor.rgb}' if fill.fgColor.type == 'rgb' else None,
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
