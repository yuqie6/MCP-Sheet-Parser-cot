"""
CSV解析器模块

解析CSV文件并转换为Sheet对象，支持UTF-8编码。
"""

import csv
from pathlib import Path
from src.models.table_model import Sheet, Row, Cell
from src.parsers.base_parser import BaseParser


class CsvParser(BaseParser):
    """CSV文件解析器，将CSV文件转换为标准化的Sheet对象。"""

    def parse(self, file_path: str) -> Sheet:
        """
        解析CSV文件并转换为Sheet对象。

        Args:
            file_path: CSV文件路径

        Returns:
            包含CSV数据的Sheet对象

        Raises:
            FileNotFoundError: 当文件不存在时
            UnicodeDecodeError: 当文件编码不是UTF-8时
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
