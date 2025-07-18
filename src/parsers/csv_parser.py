"""
CSV解析器模块

解析CSV文件并转换为Sheet对象，支持UTF-8编码和流式读取。
"""

import csv
from pathlib import Path
from collections.abc import Iterator
from src.models.table_model import Sheet, Row, Cell, LazySheet
from src.parsers.base_parser import BaseParser


class CsvRowProvider:
    """CSV文件的惰性行提供者，支持按需流式读取。"""

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self._total_rows_cache: int | None = None
        self._encoding = self._detect_encoding()

    def _detect_encoding(self) -> str:
        """检测CSV文件的编码格式。"""
        try:
            with open(self.file_path, mode='r', encoding='utf-8') as f:
                f.read(1024)  # 尝试读取一小段内容
            return 'utf-8'
        except UnicodeDecodeError:
            return 'gbk'

    def iter_rows(self, start_row: int = 0, max_rows: int | None = None) -> Iterator[Row]:
        """使用csv.reader生成器按需产出行。"""
        with open(self.file_path, mode='r', encoding=self._encoding) as csvfile:
            reader = csv.reader(csvfile)

            # 跳过到start_row
            for _ in range(start_row):
                try:
                    next(reader)
                except StopIteration:
                    return

            # 产出最多max_rows行
            count = 0
            for row_data in reader:
                if max_rows is not None and count >= max_rows:
                    break

                cells = [Cell(value=item) for item in row_data]
                yield Row(cells=cells)
                count += 1

    def get_row(self, row_index: int) -> Row:
        """按索引获取指定行。"""
        with open(self.file_path, mode='r', encoding=self._encoding) as csvfile:
            reader = csv.reader(csvfile)

            # 跳转到目标行
            for i, row_data in enumerate(reader):
                if i == row_index:
                    cells = [Cell(value=item) for item in row_data]
                    return Row(cells=cells)

            raise IndexError(f"行索引 {row_index} 超出范围")

    def get_total_rows(self) -> int:
        """无需加载全部数据即可获取总行数。"""
        if self._total_rows_cache is None:
            with open(self.file_path, mode='r', encoding=self._encoding) as csvfile:
                reader = csv.reader(csvfile)
                self._total_rows_cache = sum(1 for _ in reader)
        return self._total_rows_cache


class CsvParser(BaseParser):
    """
    CSV文件解析器。

    该解析器将CSV文件转换为标准化的Sheet对象。支持UTF-8和GBK编码，并支持大文件流式处理。
    """

    def _extract_style(self, cell):
        # CSV不支持样式，返回None
        return None

    def parse(self, file_path: str) -> list[Sheet]:
        """
        解析CSV文件并转换为Sheet对象列表。

        首先尝试使用UTF-8解码文件，如遇UnicodeDecodeError则回退为GBK。
        CSV文件只包含一个工作表，因此返回包含单个Sheet的列表。

        参数：
            file_path: CSV文件的绝对路径。

        返回：
            包含CSV数据的Sheet对象列表。

        异常：
            FileNotFoundError: 文件不存在时抛出。
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
            # 尝试使用GBK编码
            with open(path, mode='r', encoding='gbk') as csvfile:
                reader = csv.reader(csvfile)
                for row_data in reader:
                    cells = [Cell(value=item) for item in row_data]
                    rows.append(Row(cells=cells))

        sheet = Sheet(name=sheet_name, rows=rows)
        return [sheet]
    
    def supports_streaming(self) -> bool:
        """CSV解析器支持流式处理。"""
        return True
    
    def create_lazy_sheet(self, file_path: str, sheet_name: str | None = None) -> LazySheet:
        """
        创建用于流式读取CSV数据的LazySheet。

        参数：
            file_path: CSV文件的绝对路径。
            sheet_name: 工作表名称（可选，默认为文件名）。

        返回：
            可按需流式读取数据的LazySheet对象。
        """
        path = Path(file_path)
        name = sheet_name or path.stem
        provider = CsvRowProvider(file_path)
        return LazySheet(name=name, provider=provider)
