import pytest
from unittest.mock import MagicMock, PropertyMock
from src.utils.style_parser import (
    extract_style,
    extract_cell_value,
    extract_fill_color
)
from src.models.table_model import Style, RichTextFragment

@pytest.fixture
def mock_cell_factory():
    """一个工厂fixture，用于创建可配置的模拟openpyxl单元格对象。"""
    def _create_mock_cell(
        value=None,
        has_style=False,
        font=None,
        fill=None,
        alignment=None,
        border=None,
        number_format='General',
        hyperlink=None,
        comment=None
    ):
        mock = MagicMock()
        # 使用PropertyMock来正确模拟属性
        type(mock).value = PropertyMock(return_value=value)
        type(mock).has_style = PropertyMock(return_value=has_style)
        type(mock).font = PropertyMock(return_value=font)
        type(mock).fill = PropertyMock(return_value=fill)
        type(mock).alignment = PropertyMock(return_value=alignment)
        type(mock).border = PropertyMock(return_value=border)
        type(mock).number_format = PropertyMock(return_value=number_format)
        type(mock).hyperlink = PropertyMock(return_value=hyperlink)
        type(mock).comment = PropertyMock(return_value=comment)
        return mock

    return _create_mock_cell

class TestExtractCellValue:
    """测试 extract_cell_value 函数。"""

    def test_plain_text_value(self, mock_cell_factory):
        """测试提取纯文本值。"""
        cell = mock_cell_factory(value="hello")
        assert extract_cell_value(cell) == "hello"

    def test_numeric_value(self, mock_cell_factory):
        """测试提取数字值。"""
        cell = mock_cell_factory(value=123.45)
        assert extract_cell_value(cell) == 123.45

    def test_rich_text_value(self, mock_cell_factory):
        """测试提取富文本值。"""
        mock_font = MagicMock()
        mock_font.bold = True
        mock_font.italic = False
        mock_font.underline = 'single'
        mock_font.name = 'Arial'
        mock_font.size = 12
        mock_font.color = None

        rich_text_fragment = MagicMock()
        type(rich_text_fragment).text = PropertyMock(return_value="world")
        type(rich_text_fragment).font = PropertyMock(return_value=mock_font)

        cell = mock_cell_factory(value=[rich_text_fragment])
        result = extract_cell_value(cell)
        
        assert isinstance(result, list)
        assert len(result) == 1
        fragment = result[0]
        assert isinstance(fragment, RichTextFragment)
        assert fragment.text == "world"
        assert fragment.style.bold is True
        assert fragment.style.font_name == 'Arial'

class TestExtractStyle:
    """测试 extract_style 函数。"""

    def test_no_style(self, mock_cell_factory):
        """测试没有样式的单元格。"""
        cell = mock_cell_factory(has_style=False)
        style = extract_style(cell)
        assert style == Style()

    def test_font_style(self, mock_cell_factory):
        """测试提取基本的字体样式。"""
        mock_font = MagicMock()
        mock_font.bold = True
        mock_font.italic = True
        mock_font.underline = 'single'
        mock_font.size = 14
        mock_font.name = 'Calibri'
        mock_font.color = None

        cell = mock_cell_factory(has_style=True, font=mock_font)
        style = extract_style(cell)

        assert style.bold is True
        assert style.italic is True
        assert style.underline is True
        assert style.font_size == 14
        assert style.font_name == 'Calibri'

    def test_alignment_style(self, mock_cell_factory):
        """测试提取对齐样式。"""
        mock_alignment = MagicMock()
        mock_alignment.horizontal = 'center'
        mock_alignment.vertical = 'top'
        mock_alignment.wrap_text = True

        cell = mock_cell_factory(has_style=True, alignment=mock_alignment)
        style = extract_style(cell)

        assert style.text_align == 'center'
        assert style.vertical_align == 'top'
        assert style.wrap_text is True

    def test_border_style(self, mock_cell_factory):
        """测试提取边框样式。"""
        mock_border_side = MagicMock()
        mock_border_side.style = 'thin'
        mock_border_side.color = None

        mock_border = MagicMock()
        mock_border.top = mock_border_side
        mock_border.bottom = None
        mock_border.left = None
        mock_border.right = None

        cell = mock_cell_factory(has_style=True, border=mock_border)
        style = extract_style(cell)

        assert style.border_top == "1px solid #E0E0E0"
        assert style.border_bottom == ""

class TestExtractFillColor:
    """测试 extract_fill_color 函数。"""

    def test_solid_fill(self, mock_cell_factory):
        """测试提取纯色填充。"""
        mock_color = MagicMock()
        mock_color.rgb = "FFFF00"
        mock_color.theme = None
        
        mock_fill = MagicMock()
        mock_fill.patternType = 'solid'
        mock_fill.start_color = mock_color
        
        assert extract_fill_color(mock_fill) == "#FFFF00"

    def test_no_fill(self):
        """测试没有填充的情况。"""
        assert extract_fill_color(None) is None

    def test_non_solid_fill(self):
        """测试非纯色填充（如渐变）不应返回颜色。"""
        mock_fill = MagicMock()
        mock_fill.patternType = 'gradient'
        assert extract_fill_color(mock_fill) is None

class TestExtractMisc:
    """测试其他辅助提取函数。"""

    @pytest.mark.parametrize("pattern_type, expected_color", [
        ('lightGray', "#F2F2F2"),
        ('mediumGray', "#D9D9D9"),
        ('darkGray', "#BFBFBF"),
    ])
    def test_gray_pattern_fills(self, pattern_type, expected_color):
        """测试灰色系图案填充的提取。"""
        mock_fill = MagicMock()
        type(mock_fill).patternType = PropertyMock(return_value=pattern_type)
        assert extract_fill_color(mock_fill) == expected_color

    def test_gradient_fill(self):
        """测试渐变填充的提取。"""
        mock_color = MagicMock()
        type(mock_color).rgb = PropertyMock(return_value="FF0000")
        mock_stop = MagicMock()
        type(mock_stop).color = PropertyMock(return_value=mock_color)
        
        mock_fill = MagicMock()
        type(mock_fill).type = PropertyMock(return_value='gradient')
        type(mock_fill).stop = PropertyMock(return_value=[mock_stop])
        type(mock_fill).patternType = PropertyMock(return_value=None) # 确保不干扰

        assert extract_fill_color(mock_fill) == "#FF0000"

    def test_no_fill_if_pattern_is_none(self):
        """测试当patternType为None时不提取颜色。"""
        mock_fill = MagicMock()
        type(mock_fill).patternType = PropertyMock(return_value=None)
        type(mock_fill).start_color = PropertyMock(return_value=MagicMock()) # 即使有颜色信息
        assert extract_fill_color(mock_fill) is None

    def test_no_style_cell(self, mock_cell_factory):
        """测试没有样式的单元格。"""
        cell = mock_cell_factory(has_style=False)
        style = extract_style(cell)
        # 应该返回一个所有属性都为默认值的Style对象
        assert style == Style()

    def test_hyperlink_extraction(self, mock_cell_factory):
        """测试超链接的提取。"""
        mock_hyperlink = MagicMock()
        type(mock_hyperlink).target = PropertyMock(return_value="http://example.com")
        cell = mock_cell_factory(has_style=True, hyperlink=mock_hyperlink)
        style = extract_style(cell)
        assert style.hyperlink == "http://example.com"

    def test_comment_extraction(self, mock_cell_factory):
        """测试注释的提取。"""
        mock_comment = MagicMock()
        type(mock_comment).text = PropertyMock(return_value="This is a comment.")
        cell = mock_cell_factory(has_style=True, comment=mock_comment)
        style = extract_style(cell)
        assert style.comment == "This is a comment."
