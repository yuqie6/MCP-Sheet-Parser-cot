import pytest
from unittest.mock import MagicMock, PropertyMock
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

# === TDD测试：提升BorderUtils覆盖率到100% ===

def test_get_border_style_with_invalid_color(mock_border_side):
    """
    TDD测试：get_border_style应该处理无效的颜色值

    这个测试覆盖第69行的异常处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    border = mock_border_side("thin")
    border.color = MagicMock()
    border.color.rgb = "INVALID_COLOR"  # 无效颜色

    # 应该使用默认颜色而不是崩溃
    result = get_border_style(border)
    assert result == "1px solid #E0E0E0"

def test_get_border_style_with_color_attribute_error(mock_border_side):
    """
    TDD测试：get_border_style应该处理颜色属性访问错误

    这个测试覆盖第78行的异常处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    border = mock_border_side("thin")
    border.color = MagicMock()
    # 模拟访问rgb属性时抛出AttributeError
    type(border.color).rgb = PropertyMock(side_effect=AttributeError("No rgb attribute"))

    # 应该使用默认颜色
    result = get_border_style(border)
    assert result == "1px solid #E0E0E0"

def test_get_border_style_with_unknown_style(mock_border_side):
    """
    TDD测试：get_border_style应该处理未知的边框样式

    这个测试覆盖第102行的默认情况
    """
    # 🔴 红阶段：编写测试描述期望的行为
    border = mock_border_side("unknown_style", "FF0000")

    # 应该使用默认的solid样式
    result = get_border_style(border)
    assert result == "1px solid #FF0000"

def test_get_border_style_with_thick_style(mock_border_side):
    """
    TDD测试：get_border_style应该正确处理thick样式

    这个测试覆盖BORDER_STYLE_MAP中thick样式的处理
    """
    # 🔴 红阶段：编写测试描述期望的行为
    border = mock_border_side("thick", "00FF00")

    result = get_border_style(border)
    assert result == "3px solid #00FF00"

def test_parse_border_style_complete_with_complex_custom_style():
    """
    TDD测试：parse_border_style_complete应该处理复杂的自定义样式

    这个测试覆盖第113行的自定义样式处理
    """
    # 🔴 红阶段：编写测试描述期望的行为

    # 测试包含多个空格的样式
    result = parse_border_style_complete("2px   dotted", "#FF0000")
    assert result == "2px dotted #FF0000"

    # 测试已经包含颜色的样式
    result = parse_border_style_complete("1px solid red", "#00FF00")
    assert result == "1px solid red"

def test_get_xls_border_style_name_with_all_known_values():
    """
    TDD测试：get_xls_border_style_name应该处理所有已知的XLS边框样式值

    这个测试确保所有映射值都被正确处理
    """
    # 🔴 红阶段：编写测试描述期望的行为

    # 测试所有已知的XLS边框样式
    assert get_xls_border_style_name(0) == "none"
    assert get_xls_border_style_name(1) == "solid"
    assert get_xls_border_style_name(2) == "dashed"
    assert get_xls_border_style_name(3) == "dotted"
    assert get_xls_border_style_name(4) == "double"

    # 测试未知值返回默认值
    assert get_xls_border_style_name(-1) == "solid"
    assert get_xls_border_style_name(100) == "solid"

def test_format_border_color_with_edge_cases():
    """
    TDD测试：format_border_color应该处理边界情况

    这个测试覆盖各种边界情况的颜色处理
    """
    # 🔴 红阶段：编写测试描述期望的行为

    # 测试空字符串
    assert format_border_color("") == "#E0E0E0"

    # 测试不同格式的白色
    assert format_border_color("FFFFFF") == "#E0E0E0"  # 没有#前缀
    assert format_border_color("ffffff") == "#E0E0E0"  # 小写
    assert format_border_color("#ffffff") == "#E0E0E0"  # 小写带#

    # 测试接近白色的颜色
    assert format_border_color("#FEFEFE") == "#E0E0E0"
    assert format_border_color("#F0F0F0") == "#E0E0E0"

    # 测试有效的非白色颜色
    assert format_border_color("#000000") == "#000000"
    assert format_border_color("#123456") == "#123456"

def test_border_style_map_completeness():
    """
    TDD测试：验证BORDER_STYLE_MAP包含所有预期的样式

    这个测试确保边框样式映射的完整性
    """
    # 🔴 红阶段：编写测试描述期望的行为

    # 验证BORDER_STYLE_MAP包含预期的键
    expected_styles = [
        'thin', 'medium', 'thick', 'double', 'dotted', 'dashed',
        'dashDot', 'dashDotDot', 'slantDashDot', 'mediumDashed',
        'mediumDashDot', 'mediumDashDotDot'
    ]

    for style in expected_styles:
        assert style in BORDER_STYLE_MAP
        assert isinstance(BORDER_STYLE_MAP[style], str)
        assert len(BORDER_STYLE_MAP[style]) > 0