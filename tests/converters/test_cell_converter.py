import pytest
from unittest.mock import MagicMock
from datetime import datetime
from src.models.table_model import Cell, Style, RichTextFragment, RichTextFragmentStyle
from src.converters.cell_converter import CellConverter, format_chinese_date

@pytest.fixture
def mock_style_converter():
    """提供一个模拟的 StyleConverter 对象。"""
    converter = MagicMock()
    # 模拟 style_converter 的内部方法，使其返回可预测的值
    converter._format_font_family.side_effect = lambda x: x
    converter._format_font_size.side_effect = lambda x: f"{x}pt"
    return converter

@pytest.fixture
def cell_converter(mock_style_converter):
    """提供一个注入了模拟 StyleConverter 的 CellConverter 实例。"""
    return CellConverter(mock_style_converter)

@pytest.fixture
def cell_factory():
    """一个用于创建 Cell 对象的工厂 fixture。"""
    def _create_cell(value, style=None):
        return Cell(value=value, style=style or Style(), row_span=1, col_span=1)
    return _create_cell

class TestCellConverter:
    """测试 CellConverter 的核心转换逻辑。"""

    def test_convert_none_value(self, cell_converter, cell_factory):
        """测试当单元格值为 None 时，应返回空字符串。"""
        cell = cell_factory(value=None)
        assert cell_converter.convert(cell) == ""

    def test_convert_string_value(self, cell_converter, cell_factory):
        """测试纯字符串值的转换。"""
        cell = cell_factory(value="hello world")
        assert cell_converter.convert(cell) == "hello world"

    def test_convert_integer_value(self, cell_converter, cell_factory):
        """测试整数值的转换。"""
        cell = cell_factory(value=123)
        assert cell_converter.convert(cell) == "123"

    def test_convert_float_value(self, cell_converter, cell_factory):
        """测试浮点数值的转换，应保留两位小数并移除末尾多余的零。"""
        cell = cell_factory(value=123.456)
        assert cell_converter.convert(cell) == "123.46"
        
        cell_with_zero = cell_factory(value=123.40)
        assert cell_converter.convert(cell_with_zero) == "123.4"

        cell_integer_float = cell_factory(value=123.0)
        assert cell_converter.convert(cell_integer_float) == "123"
