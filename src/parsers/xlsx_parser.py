"""
XLSX解析器模块

解析Excel XLSX文件并转换为Sheet对象
包含完整的样式提取、颜色处理、边框识别等功能，支持流式读取。
"""

import openpyxl
from typing import Iterator, Optional
from src.models.table_model import Sheet, Row, Cell, LazySheet
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
            if worksheet is not None:
                self._worksheet_title_cache = worksheet.title
            else:
                self._worksheet_title_cache = ""
            workbook.close()
        return self._worksheet_title_cache
    
    def _get_merged_cells(self) -> list[str]:
        """Get merged cells info."""
        if self._merged_cells_cache is None:
            # Read merged cells from non-read-only workbook (required for merged_cells access)
            workbook = openpyxl.load_workbook(self.file_path)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            if worksheet is not None and hasattr(worksheet, "merged_cells"):
                self._merged_cells_cache = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]
            else:
                self._merged_cells_cache = []
            workbook.close()
        return self._merged_cells_cache
    
    def _parse_row(self, row_cells: tuple) -> Row:
        """Parse a tuple of openpyxl cells into a Row object."""
        cells = []
        for cell in row_cells:
            cell_value = cell.value
            cell_style = extract_style(cell) if cell else None
            cells.append(Cell(value=cell_value, style=cell_style))
        return Row(cells=cells)

    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator[Row]:
        """Yield rows on demand with complete row structure."""
        workbook = openpyxl.load_workbook(self.file_path, read_only=True)
        worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
        try:
            if worksheet is not None:
                # 获取工作表的完整尺寸
                max_row = worksheet.max_row or 0
                max_col = worksheet.max_column or 0

                # 计算实际的行范围
                end_row = max_row
                if max_rows is not None:
                    end_row = min(start_row + max_rows, max_row)

                # 使用坐标访问确保完整的行结构
                for row_idx in range(start_row + 1, end_row + 1):  # openpyxl使用1基索引
                    cells = []
                    for col_idx in range(1, max_col + 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cells.append(cell)
                    yield self._parse_row(tuple(cells))
        except Exception as e:
            # 确保在出现异常时也能正确关闭工作簿
            workbook.close()
            raise RuntimeError(f"流式读取XLSX文件失败: {str(e)}") from e
        finally:
            workbook.close()
    
    def get_row(self, row_index: int) -> Row:
        """Get a specific row by index with complete structure."""
        workbook = openpyxl.load_workbook(self.file_path, read_only=True)
        worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
        try:
            if worksheet is not None:
                max_row = worksheet.max_row or 0
                max_col = worksheet.max_column or 0

                if row_index >= max_row:
                    raise IndexError(f"Row index {row_index} out of range (max: {max_row-1})")

                # 使用坐标访问获取完整的行
                cells = []
                for col_idx in range(1, max_col + 1):
                    cell = worksheet.cell(row=row_index + 1, column=col_idx)  # openpyxl使用1基索引
                    cells.append(cell)
                return self._parse_row(tuple(cells))
            raise IndexError(f"Row index {row_index} out of range")
        except IndexError:
            # 重新抛出索引错误
            workbook.close()
            raise
        except Exception as e:
            # 处理其他异常
            workbook.close()
            raise RuntimeError(f"获取XLSX行数据失败: {str(e)}") from e
        finally:
            workbook.close()
    
    def get_total_rows(self) -> int:
        """Get total number of rows without loading all data."""
        if self._total_rows_cache is None:
            workbook = openpyxl.load_workbook(self.file_path, read_only=True)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            try:
                # Count rows using worksheet.max_row property
                if worksheet is not None and hasattr(worksheet, "max_row"):
                    self._total_rows_cache = worksheet.max_row or 0
                else:
                    self._total_rows_cache = 0
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

        # 获取工作表的实际尺寸，确保包含所有数据和样式
        max_row = worksheet.max_row or 0
        max_col = worksheet.max_column or 0

        # 性能警告：对于大型稀疏表格
        total_cells = max_row * max_col
        if total_cells > 100000:  # 10万个单元格
            import logging
            logger = logging.getLogger(__name__)
            logger.warning(f"解析大型表格 ({max_row}x{max_col}={total_cells:,}个单元格)，可能消耗较多内存")

        rows = []
        # 使用坐标访问方式确保完整的表格结构
        for row_idx in range(1, max_row + 1):
            cells = []
            for col_idx in range(1, max_col + 1):
                # 直接通过坐标访问单元格，确保包含空单元格
                cell = worksheet.cell(row=row_idx, column=col_idx)

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
