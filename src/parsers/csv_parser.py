"""
CSV解析器模块

解析CSV文件并转换为Sheet对象，支持UTF-8编码和流式读取。
"""

import csv
from pathlib import Path
from typing import Iterator, Optional
from src.models.table_model import Sheet, Row, Cell, LazySheet
from src.parsers.base_parser import BaseParser


class CsvRowProvider:
    """Lazy row provider for CSV files that streams rows on demand."""
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self._total_rows_cache: Optional[int] = None
        self._encoding = self._detect_encoding()
    
    def _detect_encoding(self) -> str:
        """Detect the encoding of the CSV file."""
        try:
            with open(self.file_path, mode='r', encoding='utf-8') as f:
                f.read(1024)  # Try reading a small chunk
            return 'utf-8'
        except UnicodeDecodeError:
            return 'gbk'
    
    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator[Row]:
        """Yield rows on demand using csv.reader as a generator."""
        with open(self.file_path, mode='r', encoding=self._encoding) as csvfile:
            reader = csv.reader(csvfile)
            
            # Skip to start_row
            for _ in range(start_row):
                try:
                    next(reader)
                except StopIteration:
                    return
            
            # Yield rows up to max_rows
            count = 0
            for row_data in reader:
                if max_rows is not None and count >= max_rows:
                    break
                
                cells = [Cell(value=item) for item in row_data]
                yield Row(cells=cells)
                count += 1
    
    def get_row(self, row_index: int) -> Row:
        """Get a specific row by index."""
        with open(self.file_path, mode='r', encoding=self._encoding) as csvfile:
            reader = csv.reader(csvfile)
            
            # Skip to the desired row
            for i, row_data in enumerate(reader):
                if i == row_index:
                    cells = [Cell(value=item) for item in row_data]
                    return Row(cells=cells)
            
            raise IndexError(f"Row index {row_index} out of range")
    
    def get_total_rows(self) -> int:
        """Get total number of rows without loading all data."""
        if self._total_rows_cache is None:
            with open(self.file_path, mode='r', encoding=self._encoding) as csvfile:
                reader = csv.reader(csvfile)
                self._total_rows_cache = sum(1 for _ in reader)
        return self._total_rows_cache


class CsvParser(BaseParser):
    """
    A parser for CSV files.

    This parser converts CSV files into a standardized Sheet object. It handles
    different encodings (UTF-8 and GBK) and supports streaming for large files.
    """

    def _extract_style(self, cell):
        # CSV不支持样式，返回None
        return None

    def parse(self, file_path: str) -> Sheet:
        """
        Parses a CSV file and converts it into a Sheet object.

        It first tries to decode the file using UTF-8, and falls back to GBK
        if a UnicodeDecodeError occurs.

        Args:
            file_path: The absolute path to the CSV file.

        Returns:
            A Sheet object containing the data from the CSV file.

        Raises:
            FileNotFoundError: If the file does not exist.
        """
        path = Path(file_path)
        sheet_name = path.stem
        rows = []

        try:
            with open(path, mode='r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row_data in reader:
                    cells = [Cell(value=item) for item in row_data]
                    rows.append(Row(cells=cells))
        except UnicodeDecodeError:
            # 尝试使用GBK编码（中文Windows常用）
            with open(path, mode='r', encoding='gbk') as csvfile:
                reader = csv.reader(csvfile)
                for row_data in reader:
                    cells = [Cell(value=item) for item in row_data]
                    rows.append(Row(cells=cells))

        return Sheet(name=sheet_name, rows=rows)
    
    def supports_streaming(self) -> bool:
        """CSV parser supports streaming."""
        return True
    
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> LazySheet:
        """
        Creates a LazySheet for streaming data from a CSV file.

        Args:
            file_path: The absolute path to the CSV file.
            sheet_name: The name of the sheet (optional, defaults to filename).

        Returns:
            A LazySheet object that can stream data on demand.
        """
        path = Path(file_path)
        name = sheet_name or path.stem
        provider = CsvRowProvider(file_path)
        return LazySheet(name=name, provider=provider)
