"""
表格服务模块

提供表格文件解析和HTML转换的核心业务逻辑。
"""

from src.parsers.factory import ParserFactory
from src.converters.html_converter import HTMLConverter


class SheetService:
    """表格服务类，协调解析器和转换器完成文件处理。"""

    def __init__(self, parser_factory: ParserFactory, html_converter: HTMLConverter):
        """
        初始化表格服务。

        Args:
            parser_factory: 解析器工厂，用于获取对应格式的解析器
            html_converter: HTML转换器，用于将表格数据转换为HTML
        """
        self.parser_factory = parser_factory
        self.html_converter = html_converter

    def convert_to_html(self, file_path: str) -> str:
        """
        将表格文件转换为HTML格式。

        Args:
            file_path: 表格文件路径

        Returns:
            HTML格式的表格内容
        """
        file_extension = file_path.split('.')[-1]
        parser = self.parser_factory.get_parser(file_extension)
        sheet_model = parser.parse(file_path)
        html_content = self.html_converter.convert(sheet_model)
        return html_content