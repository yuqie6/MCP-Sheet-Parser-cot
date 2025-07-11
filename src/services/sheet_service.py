from src.parsers.factory import ParserFactory
from src.converters.html_converter import HTMLConverter

class SheetService:
    def __init__(self, parser_factory: ParserFactory, html_converter: HTMLConverter):
        self.parser_factory = parser_factory
        self.html_converter = html_converter

    def convert_to_html(self, file_path: str) -> str:
        file_extension = file_path.split('.')[-1]
        parser = self.parser_factory.get_parser(file_extension)
        sheet_model = parser.parse(file_path)
        html_content = self.html_converter.convert(sheet_model)
        return html_content