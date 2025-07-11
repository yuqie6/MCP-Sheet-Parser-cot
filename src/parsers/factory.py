from .base_parser import BaseParser
from .xlsx_parser import XlsxParser
from .csv_parser import CsvParser
# This should be a custom exception, but for now, we'll use a standard one.
class UnsupportedFileType(Exception):
    pass

class ParserFactory:
    _parsers = {
        "csv": CsvParser(),
        "xlsx": XlsxParser(),
    }

    @staticmethod
    def get_parser(file_path: str) -> BaseParser:
        file_extension = file_path.split('.')[-1].lower()
        parser = ParserFactory._parsers.get(file_extension)
        if not parser:
            raise UnsupportedFileType(f"Unsupported file type: '{file_extension}'")
        return parser
