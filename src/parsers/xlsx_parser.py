"""
XLSX解析器模块

解析Excel XLSX文件并转换为Sheet对象
包含完整的样式提取、颜色处理、边框识别等功能，支持流式读取。
"""

import openpyxl
from typing import Iterator, Optional
from src.models.table_model import Sheet, Row, Cell, Style, LazySheet, LazyRowProvider
from src.parsers.base_parser import BaseParser
from src.utils.style_parser import extract_style


class XlsxRowProvider:
    """Lazy row provider for XLSX files using openpyxl streaming with read_only=True."""
    
    def __init__(self, file_path: str, sheet_name: Optional[str] = None):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self._total_rows_cache: Optional[int] = None
        self._merged_cells_cache: Optional[list[str]] = None
        self._worksheet_title_cache: Optional[str] = None
    
    def _get_worksheet_info(self):
        """Get worksheet info without reading all data."""
        if self._worksheet_title_cache is None:
            workbook = openpyxl.load_workbook(self.file_path, read_only=True)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            self._worksheet_title_cache = worksheet.title
            workbook.close()
        return self._worksheet_title_cache
    
    def _get_merged_cells(self) -> list[str]:
        """Get merged cells info."""
        if self._merged_cells_cache is None:
            # Read merged cells from non-read-only workbook (required for merged_cells access)
            workbook = openpyxl.load_workbook(self.file_path)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            self._merged_cells_cache = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]
            workbook.close()
        return self._merged_cells_cache
    
    def _parse_row(self, row_cells: tuple) -> Row:
        """Parse a tuple of openpyxl cells into a Row object."""
        parser = XlsxParser()
        cells = []
        for cell in row_cells:
            cell_value = cell.value
            cell_style = extract_style(cell) if cell else None
            cells.append(Cell(value=cell_value, style=cell_style))
        return Row(cells=cells)

    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator[Row]:
        """Yield rows on demand using openpyxl.iter_rows with read_only=True."""
        workbook = openpyxl.load_workbook(self.file_path, read_only=True)
        worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
        
        try:
            # Use openpyxl's iter_rows with values_only=False to get cell objects with styles
            row_count = 0
            for row_cells in worksheet.iter_rows(values_only=False):
                # Skip rows before start_row
                if row_count < start_row:
                    row_count += 1
                    continue
                
                # Stop if we've reached max_rows
                if max_rows is not None and (row_count - start_row) >= max_rows:
                    break
                
                yield self._parse_row(row_cells)
                row_count += 1
        
        finally:
            workbook.close()
    
    def get_row(self, row_index: int) -> Row:
        """Get a specific row by index."""
        workbook = openpyxl.load_workbook(self.file_path, read_only=True)
        worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
        
        try:
            # Get specific row using worksheet.iter_rows
            for i, row_cells in enumerate(worksheet.iter_rows(values_only=False)):
                if i == row_index:
                    return self._parse_row(row_cells)
            
            raise IndexError(f"Row index {row_index} out of range")
        
        finally:
            workbook.close()
    
    def get_total_rows(self) -> int:
        """Get total number of rows without loading all data."""
        if self._total_rows_cache is None:
            workbook = openpyxl.load_workbook(self.file_path, read_only=True)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            
            try:
                # Count rows using worksheet.max_row property
                self._total_rows_cache = worksheet.max_row or 0
            finally:
                workbook.close()
        
        return self._total_rows_cache


class XlsxParser(BaseParser):
    """
    A parser for XLSX files with comprehensive style extraction.

    This parser handles modern Excel files (.xlsx) and is capable of extracting
    a wide range of styling information, including fonts, colors, borders,
    and number formats. It also supports streaming for large files via the
    XlsxRowProvider.
    """

    def parse(self, file_path: str) -> Sheet:
        """
        Parses an XLSX file and returns a Sheet object.

        This method loads the entire file into memory to parse its content
        and styles. For large files, consider using `create_lazy_sheet`.

        Args:
            file_path: The absolute path to the XLSX file.

        Returns:
            A Sheet object containing the full data and styles.

        Raises:
            ValueError: If the workbook contains no active worksheet.
            FileNotFoundError: If the file does not exist.
        """
        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active

        if worksheet is None:
            raise ValueError("工作簿不包含任何活动工作表")

        rows = []
        for row in worksheet.iter_rows():
            cells = []
            for cell in row:
                # 提取单元格值和样式
                cell_value = cell.value
                cell_style = extract_style(cell)

                # 创建包含样式的Cell对象
                parsed_cell = Cell(
                    value=cell_value,
                    style=cell_style
                )
                cells.append(parsed_cell)
            rows.append(Row(cells=cells))

        merged_cells = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]

        return Sheet(
            name=worksheet.title,
            rows=rows,
            merged_cells=merged_cells
        )

    
    def supports_streaming(self) -> bool:
        """XLSX parser supports streaming."""
        return True
    
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> LazySheet:
        """
        Creates a LazySheet for streaming data from an XLSX file.

        Args:
            file_path: The absolute path to the XLSX file.
            sheet_name: The name of the sheet to parse (optional).

        Returns:
            A LazySheet object that can stream data on demand.
        """
        provider = XlsxRowProvider(file_path, sheet_name)
        name = provider._get_worksheet_info()
        merged_cells = provider._get_merged_cells()
        return LazySheet(name=name, provider=provider, merged_cells=merged_cells)
