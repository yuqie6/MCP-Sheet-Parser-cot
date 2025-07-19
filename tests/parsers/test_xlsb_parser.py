
import pytest
from unittest.mock import MagicMock, patch, mock_open
from src.parsers.xlsb_parser import XlsbParser
from src.models.table_model import Sheet, Style

@pytest.fixture
def mock_xlsb_workbook():
    """Fixture for a mocked pyxlsb workbook."""
    workbook = MagicMock()
    workbook.sheets = ["Sheet1"]
    worksheet = MagicMock()
    row = MagicMock()
    cell = MagicMock()
    cell.c = 0
    cell.v = "Test"
    row.cells = [cell]
    worksheet.rows.return_value = [[cell]]
    workbook.get_sheet.return_value.__enter__.return_value = worksheet
    return workbook

@patch('src.parsers.xlsb_parser.open_workbook')
def test_parse_success(mock_open_workbook, mock_xlsb_workbook):
    """Test successful parsing of an XLSB file."""
    mock_open_workbook.return_value.__enter__.return_value = mock_xlsb_workbook
    parser = XlsbParser()
    sheets = parser.parse("dummy.xlsb")
    assert len(sheets) == 1
    assert isinstance(sheets[0], Sheet)
    assert sheets[0].name == "Sheet1"
    assert len(sheets[0].rows) == 1
    assert len(sheets[0].rows[0].cells) == 1
    assert sheets[0].rows[0].cells[0].value == "Test"

@patch('src.parsers.xlsb_parser.open_workbook')
def test_parse_no_sheets(mock_open_workbook, mock_xlsb_workbook):
    """Test parsing a workbook with no sheets."""
    mock_xlsb_workbook.sheets = []
    mock_open_workbook.return_value.__enter__.return_value = mock_xlsb_workbook
    parser = XlsbParser()
    with pytest.raises(RuntimeError, match="工作簿不包含任何工作表"):
        parser.parse("dummy.xlsb")

def test_process_cell_value():
    """Test _process_cell_value with different value types."""
    parser = XlsbParser()
    assert parser._process_cell_value(None) is None
    assert parser._process_cell_value("Hello") == "Hello"
    assert parser._process_cell_value(123) == 123
    assert parser._process_cell_value(123.45) == 123.45
    assert parser._process_cell_value(True) is True

def test_extract_basic_style():
    """Test _extract_basic_style."""
    parser = XlsbParser()
    cell_data = MagicMock()
    cell_data.s = 1
    cell_data.f = "0.00"
    cell_data.v = 12.34
    style = parser._extract_basic_style(cell_data)
    assert isinstance(style, Style)
    assert style.number_format == "0.00"

def test_get_sheet_names(mock_xlsb_workbook):
    """Test _get_sheet_names."""
    parser = XlsbParser()
    names = parser._get_sheet_names(mock_xlsb_workbook)
    assert names == ["Sheet1"]

def test_normalize_row_data():
    """Test _normalize_row_data."""
    parser = XlsbParser()
    cell1 = MagicMock(); cell1.c = 0; cell1.v = 'A'
    cell2 = MagicMock(); cell2.c = 2; cell2.v = 'C'
    row_data = [cell1, cell2]
    normalized = parser._normalize_row_data(row_data, 3)
    assert normalized == ['A', None, 'C']

def test_streaming_support():
    """Test streaming support methods."""
    parser = XlsbParser()
    assert parser.supports_streaming() is False
    assert parser.create_lazy_sheet("dummy.xlsb") is None
