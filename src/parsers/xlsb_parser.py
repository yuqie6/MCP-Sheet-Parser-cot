import pyxlsb
from typing import Optional

from src.models.table_model import Sheet, Row, Cell, Style
from src.parsers.base_parser import BaseParser


class XlsbParser(BaseParser):
    """Parses .xlsb files using pyxlsb."""

    def parse(self, file_path: str, sheet_name: Optional[str] = None) -> Sheet:
        """Parses the given .xlsb file and returns a Sheet object."""
        with pyxlsb.open_workbook(file_path) as workbook:
            # Get sheet names
            sheet_names = workbook.get_sheet_names()
            
            # Select the worksheet
            if sheet_name:
                if sheet_name not in sheet_names:
                    raise ValueError(f"Sheet '{sheet_name}' not found in the workbook")
                target_sheet = sheet_name
            else:
                target_sheet = sheet_names[0] if sheet_names else None
                
            if not target_sheet:
                raise ValueError("No sheets found in the workbook")

            sheet_data = Sheet(name=target_sheet)
            
            # Read the worksheet data
            with workbook.get_sheet(target_sheet) as worksheet:
                rows_data = []
                for row in worksheet.rows():
                    rows_data.append(list(row))
                
                # Process the data
                for r_idx, row_values in enumerate(rows_data):
                    row_data = Row()
                    for c_idx, cell_value in enumerate(row_values):
                        # XLSB format has limited style information available through pyxlsb
                        # We create a basic style object
                        style = Style()
                        
                        # Handle different cell value types
                        if cell_value is None:
                            processed_value = None
                        elif isinstance(cell_value, (int, float, str, bool)):
                            processed_value = cell_value
                        else:
                            # Convert other types to string
                            processed_value = str(cell_value)
                        
                        cell_data = Cell(
                            value=processed_value,
                            row=r_idx,
                            column=c_idx,
                            style=style
                        )
                        row_data.cells.append(cell_data)
                    sheet_data.rows.append(row_data)
        
        return sheet_data
