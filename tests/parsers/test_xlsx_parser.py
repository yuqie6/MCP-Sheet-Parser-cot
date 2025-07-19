
import pytest
from unittest.mock import MagicMock, patch, mock_open, PropertyMock
from src.parsers.xlsx_parser import XlsxParser, XlsxRowProvider
from src.models.table_model import Sheet, LazySheet, Chart

@pytest.fixture
def mock_openpyxl_workbook():
    """Fixture for a mocked openpyxl workbook."""
    workbook = MagicMock()
    worksheet = MagicMock()
    worksheet.title = "TestSheet"
    worksheet.max_row = 1
    worksheet.max_column = 1
    cell = MagicMock()
    cell.value = "Test Data"
    cell.data_type = 's'
    # Add style attributes to the mock cell to prevent errors in extract_style
    cell.font = MagicMock()
    cell.fill = MagicMock()
    cell.border = MagicMock()
    cell.alignment = MagicMock()
    cell.hyperlink = None
    cell.comment = None
    cell.has_style = True
    worksheet.cell.return_value = cell
    worksheet.merged_cells.ranges = []

    # Mock column and row dimensions to avoid KeyError
    col_dim = MagicMock()
    col_dim.width = 10
    worksheet.column_dimensions = {'A': col_dim}
    row_dim = MagicMock()
    row_dim.height = 20
    worksheet.row_dimensions = {1: row_dim}

    worksheet.sheet_format.defaultColWidth = 8.43
    worksheet.sheet_format.defaultRowHeight = 18.0
    worksheet._charts = []
    worksheet._images = []
    workbook.sheetnames = ["TestSheet"]
    workbook.__getitem__.return_value = worksheet
    
    # Mock the 'active' property for the row provider test
    type(workbook).active = PropertyMock(return_value=worksheet)
    
    return workbook

@patch('openpyxl.load_workbook')
def test_xlsx_parser_parse(mock_load_workbook, mock_openpyxl_workbook):
    """Test basic parsing of an XLSX file."""
    mock_load_workbook.return_value = mock_openpyxl_workbook
    parser = XlsxParser()
    sheets = parser.parse("dummy.xlsx")
    assert len(sheets) == 1
    assert isinstance(sheets[0], Sheet)
    assert sheets[0].name == "TestSheet"
    assert sheets[0].rows[0].cells[0].value == "Test Data"

@patch('openpyxl.load_workbook')
def test_xlsx_row_provider_iter_rows(mock_load_workbook, mock_openpyxl_workbook):
    """Test XlsxRowProvider iter_rows method."""
    mock_load_workbook.return_value = mock_openpyxl_workbook
    provider = XlsxRowProvider("dummy.xlsx")
    rows = list(provider.iter_rows())
    assert len(rows) == 1
    assert rows[0].cells[0].value == "Test Data"

def test_xlsx_parser_streaming_support():
    """Test streaming support declaration."""
    parser = XlsxParser()
    assert parser.supports_streaming() is True

@patch('src.parsers.xlsx_parser.XlsxRowProvider')
def test_create_lazy_sheet(MockXlsxRowProvider):
    """Test create_lazy_sheet for XLSX."""
    mock_provider_instance = MockXlsxRowProvider.return_value
    mock_provider_instance._get_worksheet_info.return_value = "LazySheet"
    mock_provider_instance._get_merged_cells.return_value = []

    parser = XlsxParser()
    lazy_sheet = parser.create_lazy_sheet("dummy.xlsx")

    assert isinstance(lazy_sheet, LazySheet)
    assert lazy_sheet.name == "LazySheet"
    MockXlsxRowProvider.assert_called_once_with("dummy.xlsx", None)

@patch('openpyxl.load_workbook')
def test_extract_charts(mock_load_workbook, mock_openpyxl_workbook):
    """Test chart extraction."""
    chart = MagicMock()
    chart.title = "My Chart"
    mock_openpyxl_workbook.__getitem__.return_value._charts = [chart]
    mock_load_workbook.return_value = mock_openpyxl_workbook
    parser = XlsxParser()
    with patch.object(parser.chart_extractor, 'extract_axis_title', return_value="My Chart"):
        sheets = parser.parse("dummy.xlsx")
        assert len(sheets[0].charts) == 1
        assert sheets[0].charts[0].name == "My Chart"
