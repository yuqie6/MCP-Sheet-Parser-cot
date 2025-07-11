import csv
from pathlib import Path
from src.models.table_model import Sheet, Row, Cell
from src.parsers.base_parser import BaseParser

class CsvParser(BaseParser):
    def parse(self, file_path: str) -> Sheet:
        """
        Parses a CSV file and converts it into a Sheet object.
        The first row is assumed to be the header.
        """
        path = Path(file_path)
        sheet_name = path.stem
        rows = []
        
        with open(path, mode='r', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row_data in reader:
                cells = [Cell(value=item) for item in row_data]
                rows.append(Row(cells=cells))
        
        return Sheet(name=sheet_name, rows=rows)
