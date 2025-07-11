"""
解析器工厂模块

提供统一的解析器获取接口，支持多种表格文件格式。
"""

from .base_parser import BaseParser
from .xlsx_parser import XlsxParser
from .csv_parser import CsvParser
from .xls_parser import XlsParser
from .xlsb_parser import XlsbParser


class UnsupportedFileType(Exception):
    """不支持的文件类型异常。"""
    pass


class ParserFactory:
    """解析器工厂类，根据文件扩展名返回对应的解析器。"""

    _parsers = {
        "csv": CsvParser(),
        "xlsx": XlsxParser(),
        "xls": XlsParser(),
        "xlsb": XlsbParser(),
    }

    @staticmethod
    def get_parser(file_path: str) -> BaseParser:
        """
        根据文件路径获取对应的解析器。

        Args:
            file_path: 文件路径

        Returns:
            对应格式的解析器实例

        Raises:
            UnsupportedFileType: 当文件格式不支持时
        """
        file_extension = file_path.split('.')[-1].lower()
        parser = ParserFactory._parsers.get(file_extension)
        if not parser:
            supported_formats = ", ".join(ParserFactory._parsers.keys())
            raise UnsupportedFileType(
                f"不支持的文件格式: '{file_extension}'. "
                f"支持的格式: {supported_formats}"
            )
        return parser

    @staticmethod
    def get_supported_formats() -> list[str]:
        """
        获取所有支持的文件格式列表。

        Returns:
            支持的文件格式列表
        """
        return list(ParserFactory._parsers.keys())
