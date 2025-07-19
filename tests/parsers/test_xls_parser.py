
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

# === TDD测试：提升XLS解析器覆盖率 ===

@patch('xlrd.open_workbook')
def test_parse_with_xldate_error(mock_open_workbook, mock_workbook):
    """
    TDD测试：parse应该处理xldate转换错误

    这个测试覆盖第133-145行的异常处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_open_workbook.return_value = mock_workbook

    # 设置一个日期类型的单元格，但会导致xldate转换错误
    sheet = mock_workbook.sheet_by_index.return_value
    cell = MagicMock(spec=xlrd.sheet.Cell)
    cell.ctype = xlrd.XL_CELL_DATE
    cell.value = 99999999  # 无效的日期值
    cell.xf_index = 0
    sheet.cell.return_value = cell

    # 模拟xldate.xldate_as_datetime抛出异常
    with patch('xlrd.xldate.xldate_as_datetime', side_effect=xlrd.xldate.XLDateError("Invalid date")):
        parser = XlsParser()
        sheets = parser.parse("dummy.xls")

        # 应该返回原始数值而不是转换后的日期
        assert sheets[0].rows[0].cells[0].value == 99999999

@patch('xlrd.open_workbook')
def test_parse_with_boolean_cell(mock_open_workbook, mock_workbook):
    """
    TDD测试：parse应该正确处理布尔类型单元格

    这个测试覆盖第150行的布尔值处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_open_workbook.return_value = mock_workbook

    sheet = mock_workbook.sheet_by_index.return_value
    cell = MagicMock(spec=xlrd.sheet.Cell)
    cell.ctype = xlrd.XL_CELL_BOOLEAN
    cell.value = 1  # True
    cell.xf_index = 0
    sheet.cell.return_value = cell

    parser = XlsParser()
    sheets = parser.parse("dummy.xls")

    # 应该转换为布尔值
    assert sheets[0].rows[0].cells[0].value is True

@patch('xlrd.open_workbook')
def test_parse_with_error_cell(mock_open_workbook, mock_workbook):
    """
    TDD测试：parse应该正确处理错误类型单元格

    这个测试覆盖第153-160行的错误值处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_open_workbook.return_value = mock_workbook

    sheet = mock_workbook.sheet_by_index.return_value
    cell = MagicMock(spec=xlrd.sheet.Cell)
    cell.ctype = xlrd.XL_CELL_ERROR
    cell.value = 0x07  # #DIV/0! 错误
    cell.xf_index = 0
    sheet.cell.return_value = cell

    parser = XlsParser()
    sheets = parser.parse("dummy.xls")

    # 应该转换为错误字符串
    assert sheets[0].rows[0].cells[0].value == "#DIV/0!"

@patch('xlrd.open_workbook')
def test_parse_with_unknown_error_code(mock_open_workbook, mock_workbook):
    """
    TDD测试：parse应该处理未知的错误代码

    这个测试覆盖第159-160行的未知错误代码处理
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_open_workbook.return_value = mock_workbook

    sheet = mock_workbook.sheet_by_index.return_value
    cell = MagicMock(spec=xlrd.sheet.Cell)
    cell.ctype = xlrd.XL_CELL_ERROR
    cell.value = 0xFF  # 未知错误代码
    cell.xf_index = 0
    sheet.cell.return_value = cell

    parser = XlsParser()
    sheets = parser.parse("dummy.xls")

    # 应该返回通用错误字符串
    assert sheets[0].rows[0].cells[0].value == "#ERROR!"

def test_extract_style_with_no_xf_list():
    """
    TDD测试：_extract_style应该处理空的xf_list

    这个测试覆盖第224-226行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    parser = XlsParser()

    workbook = MagicMock()
    workbook.xf_list = []  # 空的格式列表

    style = parser._extract_style(workbook, 0)

    # 应该返回None
    assert style is None

def test_extract_style_with_invalid_xf_index():
    """
    TDD测试：_extract_style应该处理无效的xf_index

    这个测试确保方法在索引超出范围时不会崩溃
    """
    # 🔴 红阶段：编写测试描述期望的行为
    parser = XlsParser()

    workbook = MagicMock()
    workbook.xf_list = [MagicMock()]  # 只有一个格式

    style = parser._extract_style(workbook, 10)  # 超出范围的索引

    # 应该返回None
    assert style is None
