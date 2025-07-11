import csv
from typing import Optional

from src.models.table_model import Sheet, Row, Cell
from src.parsers.base_parser import BaseParser

class CsvParser(BaseParser):
    """Parses .csv files using the built-in csv module."""

    def parse(self, file_path: str, sheet_name: Optional[str] = None) -> Sheet:
        """Parses the given .csv file and returns a Sheet object."""
        sheet_data = Sheet(name=sheet_name or 'CSV Sheet')
        with open(file_path, mode='r', encoding='utf-8') as infile:
            reader = csv.reader(infile)
            for r_idx, row_values in enumerate(reader):
                row_data = Row()
                for c_idx, value in enumerate(row_values):
                    # CSV has no style information
                    cell_data = Cell(value=value, row=r_idx, column=c_idx, style=None)
                    row_data.cells.append(cell_data)
                sheet_data.rows.append(row_data)
        return sheet_data
