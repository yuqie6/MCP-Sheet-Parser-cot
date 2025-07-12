"""
解析器工厂模块

提供统一的解析器获取接口，支持5种主要表格文件格式：
- CSV: 通用文本格式
- XLSX: 现代Excel格式
- XLS: 传统Excel格式
- XLSB: Excel二进制格式
- XLSM: Excel宏文件格式
"""

from .base_parser import BaseParser
from .xlsx_parser import XlsxParser
from .csv_parser import CsvParser
from .xls_parser import XlsParser
from .xlsb_parser import XlsbParser
from .xlsm_parser import XlsmParser


class UnsupportedFileType(Exception):
    """不支持的文件类型异常。"""
    pass


class ParserFactory:
    """
    解析器工厂类，根据文件扩展名返回对应的解析器。

    支持的格式：
    - CSV (.csv): 通用逗号分隔值文件
    - XLSX (.xlsx): Excel 2007+格式，支持完整样式提取
    - XLS (.xls): Excel 97-2003格式，支持基础样式提取
    - XLSB (.xlsb): Excel二进制格式，专注数据准确性
    - XLSM (.xlsm): Excel宏文件格式，与XLSX样式一致
    """

    _parsers = {
        "csv": CsvParser(),
        "xlsx": XlsxParser(),
        "xls": XlsParser(),
        "xlsb": XlsbParser(),
        "xlsm": XlsmParser(),
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

    @staticmethod
    def get_format_info() -> dict[str, dict]:
        """
        获取所有支持格式的详细信息。

        Returns:
            格式信息字典，包含每种格式的描述和特性
        """
        return {
            "csv": {
                "name": "CSV",
                "description": "通用逗号分隔值文件",
                "features": ["数据提取", "基础格式"],
                "parser_class": "CsvParser"
            },
            "xlsx": {
                "name": "Excel XLSX",
                "description": "Excel 2007+格式",
                "features": ["完整样式提取", "95%保真度", "超链接", "注释", "合并单元格"],
                "parser_class": "XlsxParser"
            },
            "xls": {
                "name": "Excel XLS",
                "description": "Excel 97-2003格式",
                "features": ["基础样式提取", "动态颜色获取", "合并单元格"],
                "parser_class": "XlsParser"
            },
            "xlsb": {
                "name": "Excel XLSB",
                "description": "Excel二进制格式",
                "features": ["数据准确性", "基础样式", "高性能"],
                "parser_class": "XlsbParser"
            },
            "xlsm": {
                "name": "Excel XLSM",
                "description": "Excel宏文件格式",
                "features": ["完整样式提取", "宏信息保留", "95%保真度"],
                "parser_class": "XlsmParser"
            }
        }

    @staticmethod
    def is_supported_format(file_path: str) -> bool:
        """
        检查文件格式是否受支持。

        Args:
            file_path: 文件路径

        Returns:
            如果格式受支持则返回True，否则返回False
        """
        try:
            file_extension = file_path.split('.')[-1].lower()
            return file_extension in ParserFactory._parsers
        except:
            return False
