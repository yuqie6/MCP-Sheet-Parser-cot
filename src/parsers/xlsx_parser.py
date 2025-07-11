import openpyxl
from src.models.table_model import Sheet, Row, Cell
from src.parsers.base_parser import BaseParser

class XlsxParser(BaseParser):
    def parse(self, file_path: str) -> Sheet:
        """
        Parses a .xlsx file and returns a Sheet object.
        """
        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active

        if worksheet is None:
            raise ValueError("The workbook does not contain any active worksheets.")
        
        rows = []
        for row_idx, row in enumerate(worksheet.iter_rows()):
            cells = [Cell(value=cell.value) for cell in row]
            rows.append(Row(cells=cells))
            
        merged_cells = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]

        return Sheet(
            name=worksheet.title,
            rows=rows,
            merged_cells=merged_cells
        )