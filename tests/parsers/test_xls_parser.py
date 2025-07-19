
import pytest
from unittest.mock import MagicMock, patch
import xlrd
from src.parsers.xls_parser import XlsParser
from src.models.table_model import Sheet, Cell, Style

@pytest.fixture
def mock_workbook():
    """Fixture for a mocked xlrd workbook."""
    workbook = MagicMock(spec=xlrd.Book)
    workbook.nsheets = 1
    workbook.xf_list = []
    workbook.font_list = []
    workbook.format_map = {}
    workbook.colour_map = {}
    sheet = MagicMock(spec=xlrd.sheet.Sheet)
    sheet.name = "Test Sheet"
    sheet.nrows = 1
    sheet.ncols = 1
    sheet.merged_cells = []
    cell = MagicMock(spec=xlrd.sheet.Cell)
    cell.ctype = xlrd.XL_CELL_TEXT
    cell.value = "Test"
    cell.xf_index = 0
    sheet.cell.return_value = cell
    workbook.sheet_by_index.return_value = sheet
    return workbook

@patch('xlrd.open_workbook')
def test_parse_success(mock_open_workbook, mock_workbook):
    """Test successful parsing of an XLS file."""
    mock_open_workbook.return_value = mock_workbook
    parser = XlsParser()
    sheets = parser.parse("dummy.xls")
    assert len(sheets) == 1
    assert isinstance(sheets[0], Sheet)
    assert sheets[0].name == "Test Sheet"
    assert len(sheets[0].rows) == 1
    assert len(sheets[0].rows[0].cells) == 1
    assert sheets[0].rows[0].cells[0].value == "Test"

@patch('xlrd.open_workbook')
def test_parse_no_sheets(mock_open_workbook, mock_workbook):
    """Test parsing a workbook with no sheets."""
    mock_workbook.nsheets = 0
    mock_open_workbook.return_value = mock_workbook
    parser = XlsParser()
    with pytest.raises(RuntimeError, match="工作簿不包含任何工作表"):
        parser.parse("dummy.xls")

def test_get_cell_value_types(mock_workbook):
    """Test _get_cell_value with different cell types."""
    parser = XlsParser()
    sheet = mock_workbook.sheet_by_index(0)
    # Text
    sheet.cell(0, 0).ctype = xlrd.XL_CELL_TEXT
    sheet.cell(0, 0).value = "Hello"
    assert parser._get_cell_value(mock_workbook, sheet, 0, 0) == "Hello"
    # Number
    sheet.cell(0, 0).ctype = xlrd.XL_CELL_NUMBER
    sheet.cell(0, 0).value = 123.0
    assert parser._get_cell_value(mock_workbook, sheet, 0, 0) == 123
    # Date
    sheet.cell(0, 0).ctype = xlrd.XL_CELL_DATE
    sheet.cell(0, 0).value = 44197.0 # 2020-12-31
    # Boolean
    sheet.cell(0, 0).ctype = xlrd.XL_CELL_BOOLEAN
    sheet.cell(0, 0).value = 1
    assert parser._get_cell_value(mock_workbook, sheet, 0, 0) is True

def test_extract_style(mock_workbook):
    """Test _extract_style."""
    parser = XlsParser()
    sheet = mock_workbook.sheet_by_index(0)
    # Setup mock style info
    font = MagicMock()
    font.bold = 1
    font.italic = 0
    font.underline_type = 0
    font.height = 240 # 12pt
    font.name = "Arial"
    font.colour_index = 1 # White
    mock_workbook.font_list = [font]
    xf = MagicMock()
    xf.font_index = 0
    mock_workbook.xf_list = [xf]
    sheet.cell(0, 0).xf_index = 0

    style = parser._extract_style(mock_workbook, sheet, 0, 0)
    assert isinstance(style, Style)
    assert style.bold is True
    assert style.font_size == 12.0
    assert style.font_name == "Arial"

def test_get_color_from_index(mock_workbook):
    """Test _get_color_from_index."""
    parser = XlsParser()
    # From workbook colour_map
    mock_workbook.colour_map = {8: (255, 0, 0)} # Red
    assert parser._get_color_from_index(mock_workbook, 8) == "#FF0000"
    # From default_color_map
    assert parser._get_color_from_index(mock_workbook, 2) == "#FF0000"
    # Unknown index
    assert parser._get_color_from_index(mock_workbook, 999) == "#000000"

def test_extract_merged_cells(mock_workbook):
    """Test _extract_merged_cells."""
    parser = XlsParser()
    sheet = mock_workbook.sheet_by_index(0)
    sheet.merged_cells = [(0, 2, 0, 2)] # A1:B2
    merged = parser._extract_merged_cells(sheet)
    assert len(merged) == 1
    assert merged[0] == "A1:B2"

def test_index_to_excel_cell():
    """Test _index_to_excel_cell."""
    parser = XlsParser()
    assert parser._index_to_excel_cell(0, 0) == "A1"
    assert parser._index_to_excel_cell(9, 25) == "Z10"
    assert parser._index_to_excel_cell(0, 26) == "AA1"

def test_streaming_support():
    """Test streaming support methods."""
    parser = XlsParser()
    assert parser.supports_streaming() is False
    assert parser.create_lazy_sheet("dummy.xls") is None
