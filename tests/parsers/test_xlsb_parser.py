
import pytest
from unittest.mock import MagicMock, patch, mock_open, PropertyMock
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

class TestProcessCellValueEdgeCases:
    """测试_process_cell_value的边界情况。"""

    def test_process_cell_value_date_conversion_exception(self):
        """
        TDD测试：_process_cell_value应该处理日期转换异常

        这个测试覆盖第129-132行的异常处理代码
        """
        parser = XlsbParser()

        # 模拟一个会导致日期转换异常的数值
        with patch('src.parsers.xlsb_parser.datetime') as mock_datetime:
            mock_datetime.fromordinal.side_effect = ValueError("Invalid date")

            # 测试一个在日期范围内但转换失败的数值
            result = parser._process_cell_value(40000.0)  # 在日期范围内

            # 应该返回原始数值而不是日期
            assert result == 40000.0

    def test_process_cell_value_integer_conversion(self):
        """
        TDD测试：_process_cell_value应该正确处理整数转换

        这个测试覆盖第135-136行的整数转换代码
        """
        parser = XlsbParser()

        # 测试浮点数转整数
        result = parser._process_cell_value(123.0)
        assert result == 123
        assert isinstance(result, int)

        # 测试非整数浮点数
        result = parser._process_cell_value(123.45)
        assert result == 123.45
        assert isinstance(result, float)

    def test_process_cell_value_boolean_handling(self):
        """
        TDD测试：_process_cell_value应该正确处理布尔值

        这个测试覆盖第138-139行的布尔值处理代码
        """
        parser = XlsbParser()

        # 测试布尔值
        assert parser._process_cell_value(True) is True
        assert parser._process_cell_value(False) is False

    def test_process_cell_value_other_types_conversion(self):
        """
        TDD测试：_process_cell_value应该将其他类型转换为字符串

        这个测试覆盖第140-142行的其他类型转换代码
        """
        parser = XlsbParser()

        # 测试列表转字符串
        result = parser._process_cell_value([1, 2, 3])
        assert result == "[1, 2, 3]"

        # 测试字典转字符串
        result = parser._process_cell_value({"key": "value"})
        assert result == "{'key': 'value'}"

class TestExtractBasicStyleEdgeCases:
    """测试_extract_basic_style的边界情况。"""

    def test_extract_basic_style_empty_cell_data(self):
        """
        TDD测试：_extract_basic_style应该处理空单元格数据

        这个测试覆盖第151行的空数据处理代码
        """
        parser = XlsbParser()

        # 测试None
        result = parser._extract_basic_style(None)
        assert result is None

        # 测试空字典
        result = parser._extract_basic_style({})
        assert result is None

        # 测试空列表
        result = parser._extract_basic_style([])
        assert result is None

    def test_extract_basic_style_date_number_format(self):
        """
        TDD测试：_extract_basic_style应该为日期设置正确的数字格式

        这个测试覆盖第182行的日期格式设置代码
        """
        parser = XlsbParser()

        # 创建模拟的单元格数据，包含日期范围内的数值
        mock_cell_data = MagicMock()
        mock_cell_data.v = 44000.0  # 在日期范围内的数值
        mock_cell_data.s = 1  # 样式索引，表示有自定义格式

        result = parser._extract_basic_style(mock_cell_data)

        # 验证返回了Style对象且设置了日期格式
        assert isinstance(result, Style)
        assert result.number_format == "yyyy-mm-dd"

    def test_extract_basic_style_integer_number_format(self):
        """
        TDD测试：_extract_basic_style应该为整数设置正确的数字格式

        这个测试覆盖第185-186行的整数格式设置代码
        """
        parser = XlsbParser()

        # 创建模拟的单元格数据，包含整数
        mock_cell_data = MagicMock()
        mock_cell_data.v = 123
        mock_cell_data.s = 1  # 样式索引，表示有自定义格式

        result = parser._extract_basic_style(mock_cell_data)

        # 验证返回了Style对象且设置了整数格式
        assert isinstance(result, Style)
        assert result.number_format == "0"

class TestGetSheetNamesExceptionHandling:
    """测试get_sheet_names的异常处理。"""

    def test_get_sheet_names_exception_handling(self):
        """
        TDD测试：_get_sheet_names应该处理异常情况

        这个测试覆盖第210-212行的异常处理代码
        """
        parser = XlsbParser()

        # 创建一个会抛出异常的模拟工作簿
        mock_workbook = MagicMock()
        # 直接让sheets属性访问时抛出异常
        type(mock_workbook).sheets = PropertyMock(side_effect=RuntimeError("Access error"))

        with patch('src.parsers.xlsb_parser.logger') as mock_logger:
            result = parser._get_sheet_names(mock_workbook)

            # 验证返回空列表
            assert result == []

            # 验证记录了警告日志
            mock_logger.warning.assert_called_once()
            warning_call = mock_logger.warning.call_args[0][0]
            assert "获取工作表名称失败" in warning_call

class TestParseWithNullCellValues:
    """测试解析包含空值单元格的情况。"""

    @patch('src.parsers.xlsb_parser.open_workbook')
    def test_parse_with_null_cell_values(self, mock_open_workbook):
        """
        TDD测试：parse应该正确处理包含空值的单元格

        这个测试覆盖第71-72行的空值处理代码
        """

        # 创建包含空值单元格的模拟工作簿
        workbook = MagicMock()
        workbook.sheets = ["Sheet1"]

        worksheet = MagicMock()

        # 创建一个空值单元格
        null_cell = MagicMock()
        null_cell.c = 0
        null_cell.v = None  # 空值

        # 创建一个正常单元格
        normal_cell = MagicMock()
        normal_cell.c = 1
        normal_cell.v = "Test"

        worksheet.rows.return_value = [[null_cell, normal_cell]]
        workbook.get_sheet.return_value.__enter__.return_value = worksheet

        mock_open_workbook.return_value.__enter__.return_value = workbook

        parser = XlsbParser()
        sheets = parser.parse("dummy.xlsb")

        # 验证解析结果
        assert len(sheets) == 1
        assert len(sheets[0].rows) == 1
        assert len(sheets[0].rows[0].cells) == 2

        # 验证空值单元格被正确处理
        assert sheets[0].rows[0].cells[0].value is None
        assert sheets[0].rows[0].cells[0].style is None

        # 验证正常单元格
        assert sheets[0].rows[0].cells[1].value == "Test"
