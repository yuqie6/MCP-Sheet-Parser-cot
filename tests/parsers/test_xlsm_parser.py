
import pytest
from unittest.mock import MagicMock, patch, PropertyMock
from src.parsers.xlsm_parser import XlsmParser
from src.models.table_model import Sheet, LazySheet

@pytest.fixture
def mock_openpyxl_workbook():
    """Fixture for a mocked openpyxl workbook for XLSM."""
    workbook = MagicMock()
    worksheet = MagicMock()
    worksheet.title = "MacroSheet"
    worksheet.max_row = 1
    worksheet.max_column = 1
    cell = MagicMock()
    cell.value = "Macro Test"
    worksheet.cell.return_value = cell
    worksheet.merged_cells.ranges = []
    workbook.worksheets = [worksheet]
    # Mock VBA archive for macro-related tests
    vba_archive = MagicMock()
    vba_archive.namelist.return_value = ['vbaProject.bin', 'module1.bas']
    workbook.vba_archive = vba_archive
    return workbook

@patch('openpyxl.load_workbook')
def test_parse_calls_with_keep_vba(mock_load_workbook, mock_openpyxl_workbook):
    """Test that parse calls load_workbook with keep_vba=True."""
    mock_load_workbook.return_value = mock_openpyxl_workbook
    parser = XlsmParser()
    parser.parse("dummy.xlsm")
    mock_load_workbook.assert_called_once_with("dummy.xlsm", keep_vba=True)

@patch('openpyxl.load_workbook')
def test_get_macro_info_with_macros(mock_load_workbook, mock_openpyxl_workbook):
    """Test get_macro_info for a file with macros."""
    mock_load_workbook.return_value = mock_openpyxl_workbook
    parser = XlsmParser()
    info = parser.get_macro_info("dummy.xlsm")
    assert info["has_macros"] is True
    assert info["vba_files_count"] == 2
    assert "module1.bas" in info["vba_modules"]

@patch('openpyxl.load_workbook')
def test_get_macro_info_no_macros(mock_load_workbook, mock_openpyxl_workbook):
    """Test get_macro_info for a file without macros."""
    # Simulate no VBA archive
    type(mock_openpyxl_workbook).vba_archive = PropertyMock(return_value=None)
    mock_load_workbook.return_value = mock_openpyxl_workbook
    parser = XlsmParser()
    info = parser.get_macro_info("dummy.xlsm")
    assert info["has_macros"] is False

@patch('src.parsers.xlsm_parser.XlsmParser.get_macro_info')
def test_is_macro_enabled_file(mock_get_macro_info):
    """Test is_macro_enabled_file."""
    parser = XlsmParser()
    # Test non-xlsm file
    assert parser.is_macro_enabled_file("test.xlsx") is False
    # Test xlsm with macros
    mock_get_macro_info.return_value = {"has_macros": True}
    assert parser.is_macro_enabled_file("test.xlsm") is True
    # Test xlsm without macros
    mock_get_macro_info.return_value = {"has_macros": False}
    assert parser.is_macro_enabled_file("test.xlsm") is False

def test_streaming_support():
    """Test streaming support methods for XLSM."""
    parser = XlsmParser()
    assert parser.supports_streaming() is True

@patch('src.parsers.xlsx_parser.XlsxRowProvider')
def test_create_lazy_sheet(MockXlsxRowProvider):
    """Test create_lazy_sheet for XLSM."""
    mock_provider_instance = MockXlsxRowProvider.return_value
    mock_provider_instance._get_worksheet_info.return_value = "LazySheet"
    mock_provider_instance._get_merged_cells.return_value = []

    parser = XlsmParser()
    lazy_sheet = parser.create_lazy_sheet("dummy.xlsm")

    assert isinstance(lazy_sheet, LazySheet)
    assert lazy_sheet.name == "LazySheet"
    MockXlsxRowProvider.assert_called_once_with("dummy.xlsm", None)
