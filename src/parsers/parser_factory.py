import os

from src.parsers.base_parser import BaseParser
from src.parsers.xlsx_parser import XlsxParser
from src.parsers.csv_parser import CsvParser
from src.parsers.xls_parser import XlsParser
from src.parsers.xlsb_parser import XlsbParser
from src.parsers.xlsm_parser import XlsmParser
from src.parsers.et_parser import EtParser


def get_parser(file_path: str) -> BaseParser:
    """Factory function to get the appropriate parser for a file."""
    _, extension = os.path.splitext(file_path)
    extension = extension.lower()

    if extension == '.xlsx':
        return XlsxParser()
    elif extension == '.xls':
        return XlsParser()
    elif extension == '.xlsb':
        return XlsbParser()
    elif extension == '.xlsm':
        return XlsmParser()
    elif extension in ['.et', '.ett', '.ets']:
        return EtParser()
    elif extension == '.csv':
        return CsvParser()
    else:
        raise ValueError(f"Unsupported file type: {extension}. Supported formats: .xlsx, .xls, .xlsb, .xlsm, .et, .ett, .ets, .csv")
