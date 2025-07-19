
import pytest
from unittest.mock import patch
from src.parsers.factory import ParserFactory
from src.parsers.xlsx_parser import XlsxParser
from src.parsers.csv_parser import CsvParser
from src.exceptions import UnsupportedFileTypeError

@patch('src.parsers.factory.validate_file_input')
def test_get_parser_xlsx(mock_validate):
    """Test getting an XlsxParser."""
    mock_validate.return_value = ("dummy.xlsx", "xlsx")
    parser = ParserFactory.get_parser("dummy.xlsx")
    assert isinstance(parser, XlsxParser)

@patch('src.parsers.factory.validate_file_input')
def test_get_parser_csv(mock_validate):
    """Test getting a CsvParser."""
    mock_validate.return_value = ("dummy.csv", "csv")
    parser = ParserFactory.get_parser("dummy.csv")
    assert isinstance(parser, CsvParser)

@patch('src.parsers.factory.validate_file_input')
def test_get_parser_unsupported(mock_validate):
    """Test getting a parser for an unsupported file type."""
    mock_validate.return_value = ("dummy.txt", "txt")
    with pytest.raises(UnsupportedFileTypeError):
        ParserFactory.get_parser("dummy.txt")

def test_get_supported_formats():
    """Test getting the list of supported formats."""
    formats = ParserFactory.get_supported_formats()
    assert "xlsx" in formats
    assert "csv" in formats

def test_is_supported_format():
    """Test checking if a format is supported."""
    assert ParserFactory.is_supported_format("test.xlsx") is True
    assert ParserFactory.is_supported_format("test.txt") is False

@patch('src.parsers.factory.ParserFactory.get_parser')
def test_supports_streaming(mock_get_parser):
    """Test checking if a format supports streaming."""
    mock_parser = mock_get_parser.return_value
    mock_parser.supports_streaming.return_value = True
    assert ParserFactory.supports_streaming("streaming.xlsx") is True
    mock_parser.supports_streaming.return_value = False
    assert ParserFactory.supports_streaming("non_streaming.xls") is False
