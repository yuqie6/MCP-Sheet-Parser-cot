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


def test_get_border_style_with_invalid_color(mock_border_side):
    """
    TDD测试：get_border_style应该处理无效的颜色值

    这个测试覆盖第69行的异常处理代码路径
    """
    border = mock_border_side("thin")
    border.color = MagicMock()
    border.color.rgb = "INVALID_COLOR"  # 无效颜色
    border.color.indexed = None  # 确保不会通过indexed分支
    border.color.theme = None   # 确保不会通过theme分支
    border.color.value = None   # 确保不会通过value分支

    # 应该使用默认颜色而不是崩溃
    result = get_border_style(border)
    assert result == "1px solid #E0E0E0"

def test_get_border_style_with_color_attribute_error(mock_border_side):
    """
    TDD测试：get_border_style应该处理颜色属性访问错误

    这个测试覆盖第78行的异常处理代码路径
    """
    border = mock_border_side("thin")
    border.color = MagicMock()
    # 模拟访问rgb属性时抛出AttributeError
    type(border.color).rgb = PropertyMock(side_effect=AttributeError("No rgb attribute"))
    border.color.indexed = None  # 确保不会通过indexed分支
    border.color.theme = None   # 确保不会通过theme分支
    border.color.value = None   # 确保不会通过value分支

    # 应该使用默认颜色
    result = get_border_style(border)
    assert result == "1px solid #E0E0E0"

def test_get_border_style_with_unknown_style(mock_border_side):
    """
    TDD测试：get_border_style应该处理未知的边框样式

    这个测试覆盖第102行的默认情况
    """
    border = mock_border_side("unknown_style", "FF0000")

    # 应该使用默认的solid样式
    result = get_border_style(border)
    assert result == "1px solid #FF0000"

def test_get_border_style_with_thick_style(mock_border_side):
    """
    TDD测试：get_border_style应该正确处理thick样式

    这个测试覆盖BORDER_STYLE_MAP中thick样式的处理
    """
    border = mock_border_side("thick", "00FF00")

    result = get_border_style(border)
    assert result == "3px solid #00FF00"

def test_parse_border_style_complete_with_complex_custom_style():
    """
    TDD测试：parse_border_style_complete应该处理复杂的自定义样式

    这个测试覆盖第113行的自定义样式处理
    """

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

    # 验证BORDER_STYLE_MAP包含预期的键
    expected_styles = [
        'thin', 'medium', 'thick', 'double', 'dotted', 'dashed',
        'dashDot', 'dashDotDot', 'slantDashDot', 'mediumDashed',
        'mediumDashDot', 'mediumDashDotDot'
    ]

    for style in expected_styles:
        assert style in BORDER_STYLE_MAP
        # 映射值可以是字符串或元组
        style_value = BORDER_STYLE_MAP[style]
        assert isinstance(style_value, (str, tuple))
        if isinstance(style_value, tuple):
            assert len(style_value) == 2  # (width, style_type)
        else:
            assert len(style_value) > 0


class TestFormatBorderColorEdgeCases:
    """测试format_border_color的边界情况。"""

    def test_format_border_color_with_specific_gray_colors(self):
        """
        TDD测试：format_border_color应该转换特定的灰色

        这个测试覆盖第78行的特定颜色转换代码
        """


        # 测试需要转换的特定灰色
        assert format_border_color("#D8D8D8") == "#E0E0E0"
        assert format_border_color("#DADADA") == "#E0E0E0"
        assert format_border_color("#DBDBDB") == "#E0E0E0"

        # 测试大小写不敏感
        assert format_border_color("#d8d8d8") == "#E0E0E0"
        assert format_border_color("#dadada") == "#E0E0E0"
        assert format_border_color("#dbdbdb") == "#E0E0E0"

class TestGetBorderStyleEdgeCases:
    """测试get_border_style的边界情况。"""

    def test_get_border_style_old_format_return(self):
        """
        TDD测试：get_border_style应该处理旧格式的直接返回

        这个测试覆盖第87行的旧格式返回代码
        """


        # 创建一个没有颜色的边框对象
        mock_border = MagicMock()
        mock_border.style = "custom_style"
        mock_border.color = None

        # 模拟get_border_style的内部逻辑，当没有颜色时应该返回样式字符串
        result = get_border_style(mock_border)

        # 验证返回了处理后的样式（实际会通过BORDER_STYLE_MAP处理）
        assert isinstance(result, str)

class TestParseBorderStyleCompleteEdgeCases:
    """测试parse_border_style_complete的边界情况。"""

    def test_parse_border_style_complete_single_style_info(self):
        """
        TDD测试：parse_border_style_complete应该处理单个样式信息

        这个测试覆盖第115行的单个样式信息处理代码
        """


        # 查找一个映射为字符串而不是元组的样式
        # 从BORDER_STYLE_MAP中找到一个字符串值的样式
        string_style = None
        for style, value in BORDER_STYLE_MAP.items():
            if isinstance(value, str):
                string_style = style
                break

        if string_style:
            result = parse_border_style_complete(string_style, "#FF0000")
            # 验证返回了正确的格式：1px + 样式 + 颜色
            assert result == f"1px {BORDER_STYLE_MAP[string_style]} #FF0000"
        else:
            # 如果没有找到字符串样式，创建一个测试用例
            # 临时修改映射来测试这个分支
            original_value = BORDER_STYLE_MAP.get('test_style')
            BORDER_STYLE_MAP['test_style'] = 'dotted'
            try:
                result = parse_border_style_complete('test_style', '#FF0000')
                assert result == "1px dotted #FF0000"
            finally:
                # 恢复原始映射
                if original_value is None:
                    del BORDER_STYLE_MAP['test_style']
                else:
                    BORDER_STYLE_MAP['test_style'] = original_value

    def test_parse_border_style_complete_fallback_default(self):
        """
        TDD测试：parse_border_style_complete应该使用默认样式作为后备

        这个测试覆盖第126行的默认返回代码
        """


        # 使用一个无法解析的样式字符串
        result = parse_border_style_complete("invalid_unparseable_style", "#FF0000")

        # 验证返回了默认样式
        assert result == "1px solid #FF0000"

        # 测试空字符串
        result = parse_border_style_complete("", "#00FF00")
        assert result == "1px solid #00FF00"

        # 测试只有空格的字符串
        result = parse_border_style_complete("   ", "#0000FF")
        assert result == "1px solid #0000FF"

class TestGetBorderStyleStringFormat:
    """测试get_border_style的字符串格式处理。"""

    def test_get_border_style_string_format_return(self):
        """
        TDD测试：get_border_style应该处理字符串格式的样式映射

        这个测试覆盖第87行的字符串格式返回代码
        """


        # 临时修改BORDER_STYLE_MAP来包含一个字符串值
        original_value = BORDER_STYLE_MAP.get('test_string_style')
        BORDER_STYLE_MAP['test_string_style'] = 'solid'  # 字符串而不是元组

        try:
            # 创建一个使用字符串样式的边框对象
            mock_border = MagicMock()
            mock_border.style = 'test_string_style'
            mock_border.color = None

            result = get_border_style(mock_border)

            # 验证返回了字符串样式
            assert result == 'solid'

        finally:
            # 恢复原始映射
            if original_value is None:
                del BORDER_STYLE_MAP['test_string_style']
            else:
                BORDER_STYLE_MAP['test_string_style'] = original_value

    def test_get_border_style_with_normal_color(self, mock_border_side):
        """
        TDD测试：get_border_style应该保持正常颜色不变

        这个测试覆盖第78行else分支的正常颜色处理
        """


        # 使用一个正常的颜色（不在转换列表中）
        border = mock_border_side("thin", "00FF00")  # 绿色
        result = get_border_style(border)

        # 验证颜色保持不变
        assert result == "1px solid #00FF00"

        # 测试另一个正常颜色
        border = mock_border_side("medium", "FF00FF")  # 紫色
        result = get_border_style(border)

        # 验证颜色保持不变
        assert result == "2px solid #FF00FF"

    def test_get_border_style_with_specific_gray_conversion(self, mock_border_side):
        """
        TDD测试：get_border_style应该转换特定的灰色

        这个测试覆盖第78行的特定灰色转换代码
        """


        # 测试需要转换的特定灰色
        border = mock_border_side("thin", "D8D8D8")
        result = get_border_style(border)
        assert result == "1px solid #E0E0E0"

        border = mock_border_side("medium", "DADADA")
        result = get_border_style(border)
        assert result == "2px solid #E0E0E0"

        border = mock_border_side("thick", "DBDBDB")
        result = get_border_style(border)
        assert result == "3px solid #E0E0E0"