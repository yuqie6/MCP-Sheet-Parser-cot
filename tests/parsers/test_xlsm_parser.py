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


@patch('openpyxl.load_workbook')
def test_parse_with_empty_workbook(mock_load_workbook):
    """
    TDD测试：parse应该处理空工作簿

    这个测试覆盖第43行的空工作簿检查代码
    """

    # 模拟一个没有工作表的工作簿
    empty_workbook = MagicMock()
    empty_workbook.worksheets = []
    mock_load_workbook.return_value = empty_workbook

    parser = XlsmParser()

    # 应该抛出RuntimeError（因为ValueError被包装了）
    with pytest.raises(RuntimeError, match="无法解析XLSM文件"):
        parser.parse("empty.xlsm")

@patch('openpyxl.load_workbook')
def test_parse_with_exception(mock_load_workbook):
    """
    TDD测试：parse应该处理解析异常

    这个测试覆盖第89-91行的异常处理代码
    """

    # 模拟openpyxl抛出异常
    mock_load_workbook.side_effect = Exception("File corrupted")

    parser = XlsmParser()

    # 应该抛出RuntimeError
    with pytest.raises(RuntimeError, match="无法解析XLSM文件"):
        parser.parse("corrupted.xlsm")

@patch('openpyxl.load_workbook')
def test_log_macro_info_with_vba_exception(mock_load_workbook, mock_openpyxl_workbook):
    """
    TDD测试：_log_macro_info应该处理VBA信息获取异常

    这个测试覆盖第120-121行的VBA异常处理代码
    """

    # 模拟VBA archive存在但获取详细信息时抛出异常
    vba_archive = MagicMock()
    vba_archive.namelist.side_effect = Exception("VBA access error")
    mock_openpyxl_workbook.vba_archive = vba_archive
    mock_load_workbook.return_value = mock_openpyxl_workbook

    parser = XlsmParser()

    # 应该不抛出异常，只是记录日志
    sheets = parser.parse("test.xlsm")
    assert len(sheets) == 1

@patch('openpyxl.load_workbook')
def test_log_macro_info_without_vba(mock_load_workbook, mock_openpyxl_workbook):
    """
    TDD测试：_log_macro_info应该处理没有VBA的情况

    这个测试覆盖第122-123行的无VBA处理代码
    """

    # 模拟没有VBA archive
    type(mock_openpyxl_workbook).vba_archive = PropertyMock(return_value=None)
    mock_load_workbook.return_value = mock_openpyxl_workbook

    parser = XlsmParser()

    # 应该正常解析，不抛出异常
    sheets = parser.parse("no_vba.xlsm")
    assert len(sheets) == 1

@patch('openpyxl.load_workbook')
def test_log_macro_info_with_general_exception(mock_load_workbook, mock_openpyxl_workbook):
    """
    TDD测试：_log_macro_info应该处理一般异常

    这个测试覆盖第125-126行的一般异常处理代码
    """

    # 模拟访问vba_archive属性时抛出异常
    type(mock_openpyxl_workbook).vba_archive = PropertyMock(side_effect=Exception("General error"))
    mock_load_workbook.return_value = mock_openpyxl_workbook

    parser = XlsmParser()

    # 应该正常解析，不抛出异常
    sheets = parser.parse("error.xlsm")
    assert len(sheets) == 1

@patch('openpyxl.load_workbook')
def test_get_macro_info_with_exception(mock_load_workbook):
    """
    TDD测试：get_macro_info应该处理加载异常

    这个测试覆盖第162-163行的异常处理代码
    """

    # 模拟openpyxl抛出异常
    mock_load_workbook.side_effect = Exception("File not found")

    parser = XlsmParser()

    # 应该返回默认的宏信息
    info = parser.get_macro_info("missing.xlsm")
    assert info["has_macros"] is False
    assert info["vba_files_count"] == 0
    assert info["vba_modules"] == []

@patch('openpyxl.load_workbook')
def test_is_macro_enabled_file_with_exception(mock_load_workbook):
    """
    TDD测试：is_macro_enabled_file应该处理加载异常

    这个测试覆盖第167-168行的异常处理代码
    """

    # 模拟openpyxl抛出异常
    mock_load_workbook.side_effect = Exception("File access error")

    parser = XlsmParser()

    # 应该返回False
    result = parser.is_macro_enabled_file("error.xlsm")
    assert result is False

# 注意：第215-217行是is_macro_enabled_file方法的异常处理，
# 已经在test_is_macro_enabled_file_with_exception中覆盖。
# create_lazy_sheet方法继承自XlsxParser，不在此模块中。
