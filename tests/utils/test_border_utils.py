import pytest
from unittest.mock import MagicMock
from src.utils.border_utils import (
    get_border_style,
    parse_border_style_complete,
    get_xls_border_style_name,
    format_border_color,
    BORDER_STYLE_MAP
)

# Mock for openpyxl BorderSide object
@pytest.fixture
def mock_border_side():
    def _mock_border_side(style, color_rgb=None):
        border_side = MagicMock()
        border_side.style = style
        if color_rgb:
            border_side.color = MagicMock()
            border_side.color.rgb = color_rgb
        else:
            border_side.color = None
        return border_side
    return _mock_border_side

def test_get_border_style_with_color(mock_border_side):
    """测试 get_border_style 函数（带颜色）。"""
    border = mock_border_side("thin", "FF0000")
    assert get_border_style(border) == "1px solid #FF0000"

def test_get_border_style_no_color(mock_border_side):
    """测试 get_border_style 函数（无颜色）。"""
    border = mock_border_side("dashed", None)
    assert get_border_style(border) == "1px dashed #E0E0E0"

def test_get_border_style_white_color_gets_converted(mock_border_side):
    """测试 get_border_style 函数，白色会被转换为灰色。"""
    border = mock_border_side("medium", "FFFFFF")
    assert get_border_style(border) == "2px solid #E0E0E0"

def test_get_border_style_no_style():
    """测试 get_border_style 函数（无样式）。"""
    assert get_border_style(None) == ""
    border = MagicMock()
    border.style = None
    assert get_border_style(border) == ""

def test_parse_border_style_complete_known_styles():
    """测试 parse_border_style_complete 函数（已知样式）。"""
    assert parse_border_style_complete("thin", "#FF0000") == "1px solid #FF0000"
    assert parse_border_style_complete("double") == "3px double #E0E0E0"

def test_parse_border_style_complete_custom_styles():
    """测试 parse_border_style_complete 函数（自定义样式）。"""
    assert parse_border_style_complete("2.5pt dotted", "blue") == "2.5pt dotted blue"
    assert parse_border_style_complete("5px solid") == "5px solid #E0E0E0"

def test_parse_border_style_complete_no_style():
    """测试 parse_border_style_complete 函数（无样式）。"""
    assert parse_border_style_complete(None) == "1px solid #E0E0E0"
    assert parse_border_style_complete("") == "1px solid #E0E0E0"

def test_get_xls_border_style_name():
    """测试 get_xls_border_style_name 函数。"""
    assert get_xls_border_style_name(1) == "solid"
    assert get_xls_border_style_name(2) == "dashed"
    assert get_xls_border_style_name(99) == "solid" # Test default

def test_format_border_color():
    """测试 format_border_color 函数。"""
    assert format_border_color(None) == "#E0E0E0"
    assert format_border_color("#FFFFFF") == "#E0E0E0"
    assert format_border_color("#FFF") == "#E0E0E0"
    assert format_border_color("#D8D8D8") == "#E0E0E0"
    assert format_border_color("#FF0000") == "#FF0000"