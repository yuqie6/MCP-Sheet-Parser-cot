import os
from typing import Type

from src.parsers.base_parser import BaseParser
from src.parsers.xlsx_parser import XlsxParser
from src.parsers.csv_parser import CsvParser


def get_parser(file_path: str) -> BaseParser:
    """Factory function to get the appropriate parser for a file."""
    _, extension = os.path.splitext(file_path)
    extension = extension.lower()

    if extension == '.xlsx':
        return XlsxParser()
    elif extension == '.csv':
        return CsvParser()
    else:
        raise ValueError(f"Unsupported file type: {extension}")
