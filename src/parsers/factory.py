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
from ..exceptions import UnsupportedFileTypeError
from ..validators import validate_file_input


class ParserFactory:
    """
    解析器工厂类，根据文件扩展名创建对应的解析器。

    该工厂为给定的文件格式提供统一的解析器获取方式。支持多种表格格式，并对不支持的类型抛出错误。

    支持的格式：
    - CSV (.csv)：通用逗号分隔值文件。
    - XLSX (.xlsx)：现代Excel格式（2007+），支持完整样式。
    - XLS (.xls)：传统Excel格式（97-2003），支持基础样式。
    - XLSB (.xlsb)：Excel二进制格式，注重数据准确性。
    - XLSM (.xlsm)：Excel宏文件格式，类似XLSX。
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
        获取指定文件路径对应的解析器。

        参数：
            file_path: 文件的绝对路径。

        返回：
            继承自 BaseParser 的解析器实例。

        异常：
            UnsupportedFileTypeError: 文件格式不支持时抛出。
            ValidationError: 文件路径无效时抛出。
            FileNotFoundError: 文件不存在时抛出。
        """
        # 使用验证器验证文件输入
        validated_path, file_extension = validate_file_input(file_path)

        parser = ParserFactory._parsers.get(file_extension)
        if not parser:
            supported_formats = list(ParserFactory._parsers.keys())
            raise UnsupportedFileTypeError(file_extension, supported_formats)

        return parser

    @staticmethod
    def get_supported_formats() -> list[str]:
        """
        返回所有支持的文件格式扩展名列表。
        """
        return list(ParserFactory._parsers.keys())

    @staticmethod
    def get_format_info() -> dict[str, dict]:
        """
        获取所有支持格式的详细信息。

        返回：
            格式信息字典，包含每种格式的描述和特性
        """
        return {
            "csv": {
                "name": "CSV",
                "description": "通用逗号分隔值文件",
                "features": ["数据提取", "基础格式", "流式读取"],
                "parser_class": "CsvParser",
                "supports_streaming": True
            },
            "xlsx": {
                "name": "Excel XLSX",
                "description": "Excel 2007+格式",
                "features": ["完整样式提取",  "超链接", "注释", "合并单元格", "流式读取"],
                "parser_class": "XlsxParser",
                "supports_streaming": True
            },
            "xls": {
                "name": "Excel XLS",
                "description": "Excel 97-2003格式",
                "features": ["基础样式提取", "动态颜色获取", "合并单元格"],
                "parser_class": "XlsParser",
                "supports_streaming": False
            },
            "xlsb": {
                "name": "Excel XLSB",
                "description": "Excel二进制格式",
                "features": ["数据准确性", "基础样式", "高性能"],
                "parser_class": "XlsbParser",
                "supports_streaming": False
            },
            "xlsm": {
                "name": "Excel XLSM",
                "description": "Excel宏文件格式",
                "features": ["完整样式提取", "宏信息保留", "流式读取"],
                "parser_class": "XlsmParser",
                "supports_streaming": True
            }
        }

    @staticmethod
    def is_supported_format(file_path: str) -> bool:
        """
        检查指定路径的文件格式是否受支持。

        参数：
            file_path: 文件路径。

        返回：
            如果支持该格式返回True，否则返回False。
        """
        try:
            file_extension = file_path.split('.')[-1].lower()
            return file_extension in ParserFactory._parsers
        except IndexError:
            return False
    
    @staticmethod
    def supports_streaming(file_path: str) -> bool:
        """
        检查指定文件格式是否支持流式读取。

        参数：
            file_path: 文件路径

        返回：
            如果格式支持流式读取则返回True，否则返回False
        """
        try:
            parser = ParserFactory.get_parser(file_path)
            return parser.supports_streaming()
        except UnsupportedFileTypeError:
            return False
    
    @staticmethod
    def get_streaming_formats() -> list[str]:
        """
        获取所有支持流式读取的文件格式列表。

        返回：
            支持流式读取的文件格式列表
        """
        streaming_formats = []
        for format_ext, parser in ParserFactory._parsers.items():
            if parser.supports_streaming():
                streaming_formats.append(format_ext)
        return streaming_formats
    
    @staticmethod
    def create_lazy_sheet(file_path: str, sheet_name: str | None = None):
        """
        为支持流式读取的文件创建懒加载工作表。

        参数：
            file_path: 文件路径
            sheet_name: 工作表名称（可选）

        返回：
            LazySheet对象，如果不支持流式读取则返回None

        异常：
            UnsupportedFileTypeError: 当文件格式不支持时抛出
        """
        parser = ParserFactory.get_parser(file_path)
        return parser.create_lazy_sheet(file_path, sheet_name)
