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
    A factory class for creating file parsers based on file extensions.

    This factory provides a centralized way to get the correct parser for a
    given file format. It supports various spreadsheet formats and handles
    errors for unsupported types.

    Supported formats:
    - CSV (.csv): Comma-Separated Values.
    - XLSX (.xlsx): Modern Excel format (2007+), with full style support.
    - XLS (.xls): Legacy Excel format (97-2003), with basic style support.
    - XLSB (.xlsb): Excel binary format, focused on data accuracy.
    - XLSM (.xlsm): Excel macro-enabled format, similar to XLSX.
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
        Retrieves the appropriate parser for a given file path.

        Args:
            file_path: The absolute path to the file.

        Returns:
            An instance of a parser that inherits from BaseParser.

        Raises:
            UnsupportedFileType: If the file format is not supported.
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
        Returns a list of all supported file format extensions.
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
        Checks if the file format of a given path is supported.

        Args:
            file_path: The path to the file.

        Returns:
            True if the format is supported, False otherwise.
        """
        try:
            file_extension = file_path.split('.')[-1].lower()
            return file_extension in ParserFactory._parsers
        except:
            return False
    
    @staticmethod
    def supports_streaming(file_path: str) -> bool:
        """
        检查指定文件格式是否支持流式读取。
        
        Args:
            file_path: 文件路径
            
        Returns:
            如果格式支持流式读取则返回True，否则返回False
        """
        try:
            parser = ParserFactory.get_parser(file_path)
            return parser.supports_streaming()
        except:
            return False
    
    @staticmethod
    def get_streaming_formats() -> list[str]:
        """
        获取所有支持流式读取的文件格式列表。
        
        Returns:
            支持流式读取的文件格式列表
        """
        streaming_formats = []
        for format_ext, parser in ParserFactory._parsers.items():
            if parser.supports_streaming():
                streaming_formats.append(format_ext)
        return streaming_formats
    
    @staticmethod
    def create_lazy_sheet(file_path: str, sheet_name: str = None):
        """
        为支持流式读取的文件创建懒加载工作表。
        
        Args:
            file_path: 文件路径
            sheet_name: 工作表名称（可选）
            
        Returns:
            LazySheet对象，如果不支持流式读取则返回None
            
        Raises:
            UnsupportedFileType: 当文件格式不支持时
        """
        parser = ParserFactory.get_parser(file_path)
        return parser.create_lazy_sheet(file_path, sheet_name)
