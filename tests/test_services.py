import pytest
from src.services.sheet_service import SheetService
from src.parsers.factory import ParserFactory
from src.converters.html_converter import HTMLConverter

@pytest.fixture
def sheet_service():
    """Provides a SheetService instance for tests."""
    parser_factory = ParserFactory()
    html_converter = HTMLConverter()
    return SheetService(parser_factory=parser_factory, html_converter=html_converter)

def test_sheet_service_initialization(sheet_service):
    """
    Tests that the SheetService can be initialized with its dependencies.
    """
    assert sheet_service is not None

def test_process_file_csv(sheet_service):
    """
    Tests the full processing of a CSV file to HTML.
    """
    file_path = 'tests/data/sample.csv'
    html_output = sheet_service.convert_to_html(file_path)
    assert "<td>Name</td>" in html_output
    assert "<td>Alice</td>" in html_output
    assert "<td>30</td>" in html_output

def test_process_file_xlsx(sheet_service):
    """
    Tests the full processing of an XLSX file to HTML.
    """
    file_path = 'tests/data/sample.xlsx'
    html_output = sheet_service.convert_to_html(file_path)
    assert "<td>Name</td>" in html_output
    assert "<td>Bob</td>" in html_output
    assert "<td>25</td>" in html_output