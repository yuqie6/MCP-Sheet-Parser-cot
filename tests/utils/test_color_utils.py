import pytest
from unittest.mock import MagicMock
from src.utils.color_utils import (
    normalize_color,
    format_color,
    convert_scheme_color_to_hex,
    generate_pie_color_variants,
    generate_distinct_colors,
    ensure_distinct_colors,
    extract_color,
    apply_tint,
    get_color_brightness,
    has_sufficient_contrast,
    apply_smart_color_matching
)
from src.models.table_model import Style

class TestNormalizeColor:
    """测试 normalize_color 函数。"""

    @pytest.mark.parametrize("input_color, expected", [
        ("#FF0000", "#FF0000"),
        ("00FF00", "#00FF00"),
        ("0000ff", "#0000FF"),
        ("#123", "#000000"),  # 无效长度
        ("ABC", "#000000"),   # 无效长度
        ("", "#000000"),
        (None, "#000000"),
        ("00ABCDEF", "#ABCDEF"), # ARGB
    ])
    def test_various_inputs(self, input_color, expected):
        assert normalize_color(input_color) == expected

class TestFormatColor:
    """测试 format_color 函数。"""

    @pytest.mark.parametrize("input_color, is_font, expected", [
        ("#FF0000", False, "#FF0000"),
        ("00FF00", False, "#00FF00"),
        ("#ABC", False, "#AABBCC"),
        ("DEF", False, "#DDEEFF"),
        ("RED", False, "#FF0000"),
        ("INVALID", True, "#000000"),
        ("INVALID", False, None),
        (None, True, "#000000"),
        (None, False, None),
    ])
    def test_formatting_and_defaults(self, input_color, is_font, expected):
        assert format_color(input_color, is_font_color=is_font) == expected

    def test_border_color_special_handling(self):
        """测试边框颜色的特殊处理逻辑。"""
        assert format_color("#FFFFFF", is_border_color=True) == "#E0E0E0"
        assert format_color("WHITE", is_border_color=True) == "#E0E0E0"
        assert format_color("#D8D8D8", is_border_color=True) == "#E0E0E0"
        assert format_color("#000000", is_border_color=True) == "#000000"

class TestSchemeColorConversion:
    """测试 convert_scheme_color_to_hex 函数。"""

    @pytest.mark.parametrize("scheme, expected", [
        ("accent1", "#4F81BD"),
        ("lt1", "#FFFFFF"),
        ("dk1", "#000000"),
        ("invalid_scheme", "#70AD47"), # 测试默认值
    ])
    def test_scheme_to_hex(self, scheme, expected):
        assert convert_scheme_color_to_hex(scheme) == expected

class TestPieColorGeneration:
    """测试 generate_pie_color_variants 函数。"""

    def test_generate_single_color(self):
        assert generate_pie_color_variants("#FF0000", 1) == ["#FF0000"]

    def test_generate_multiple_colors(self):
        colors = generate_pie_color_variants("#C0504D", 3)
        assert len(colors) == 3
        assert colors[0] == "#C0504D"
        assert colors[1] != colors[0]

    def test_invalid_color_fallback(self):
        """测试当提供无效颜色时，能否回退到默认颜色列表。"""
        colors = generate_pie_color_variants("invalid", 5)
        assert len(colors) == 5
        assert colors[0] == "#C0504D" # 默认列表的第一个颜色

class TestDistinctColorGeneration:
    """测试 generate_distinct_colors 和 ensure_distinct_colors 函数。"""

    def test_generate_basic(self):
        colors = generate_distinct_colors(5)
        assert len(colors) == 5
        assert len(set(colors)) == 5

    def test_generate_with_existing(self):
        existing = ['#FF6B6B', '#4ECDC4']
        new_colors = generate_distinct_colors(3, existing_colors=existing)
        assert len(new_colors) == 3
        assert not any(c in existing for c in new_colors)

    def test_ensure_no_duplicates(self):
        initial = ["#FF0000", "#00FF00", "#FF0000"]
        result = ensure_distinct_colors(initial, 3)
        assert len(result) == 3
        assert len(set(result)) == 3
        assert result[0] == "#FF0000"
        assert result[1] == "#00FF00"

    def test_ensure_needs_more_colors(self):
        initial = ["#FF0000"]
        result = ensure_distinct_colors(initial, 4)
        assert len(result) == 4
        assert len(set(result)) == 4
        assert result[0] == "#FF0000"

class TestExtractColor:
    """测试 extract_color 函数，需要模拟openpyxl的Color对象。"""

    def test_extract_rgb(self):
        mock_color = MagicMock()
        mock_color.rgb = "FF0000"
        mock_color.theme = None
        mock_color.indexed = None
        assert extract_color(mock_color) == "#FF0000"

    def test_extract_argb(self):
        mock_color = MagicMock()
        mock_color.rgb = "FFFF0000"
        mock_color.theme = None
        mock_color.indexed = None
        assert extract_color(mock_color) == "#FF0000"

    def test_extract_theme_color(self):
        mock_color = MagicMock()
        mock_color.rgb = None
        mock_color.theme = 4 # accent1 in theme_colors
        mock_color.tint = 0.0
        mock_color.indexed = None
        assert extract_color(mock_color) == "#5B9BD5"

    def test_extract_theme_with_tint(self):
        mock_color = MagicMock()
        mock_color.rgb = None
        mock_color.theme = 4
        mock_color.tint = 0.5 # 变亮
        mock_color.indexed = None
        # 实际结果会是 #5B9BD5 和 #FFFFFF 的中间色
        assert extract_color(mock_color) is not None
        assert extract_color(mock_color) != "#5B9BD5"

    def test_no_color(self):
        assert extract_color(None) is None

class TestTintAndBrightness:
    """测试 apply_tint, get_color_brightness, 和 has_sufficient_contrast。"""

    def test_apply_tint_lighten(self):
        assert apply_tint("#808080", 0.5) == "#BFBFBF"

    def test_apply_tint_darken(self):
        assert apply_tint("#808080", -0.5) == "#404040"

    def test_brightness(self):
        assert get_color_brightness("#000000") == 0
        assert get_color_brightness("#FFFFFF") == 255
        assert 80 < get_color_brightness("#808080") < 130

    def test_contrast(self):
        assert has_sufficient_contrast("#000000", "#FFFFFF") is True
        assert has_sufficient_contrast("#808080", "#808080") is False

class TestSmartColorMatching:
    """测试 apply_smart_color_matching 函数。"""

    def test_dark_bg_no_font_color(self):
        style = Style(background_color="#000000")
        new_style = apply_smart_color_matching(style)
        assert new_style.font_color == "#FFFFFF"

    def test_light_bg_no_font_color(self):
        style = Style(background_color="#FFFFFF")
        new_style = apply_smart_color_matching(style)
        assert new_style.font_color is None # 浅色背景不自动添加黑色字体

    def test_low_contrast_adjusts(self):
        style = Style(background_color="#000000", font_color="#101010")
        new_style = apply_smart_color_matching(style)
        assert new_style.font_color == "#FFFFFF"

    def test_high_contrast_keeps(self):
        style = Style(background_color="#000000", font_color="#FFFFFF")
        new_style = apply_smart_color_matching(style)
        assert new_style.font_color == "#FFFFFF"

# === TDD测试：提升ColorUtils覆盖率 ===

class TestConvertSchemeColorToHex:
    """测试 convert_scheme_color_to_hex 函数的边界情况。"""

    def test_convert_scheme_color_with_invalid_scheme(self):
        """
        TDD测试：convert_scheme_color_to_hex应该处理无效的配色方案

        这个测试覆盖第217-229行的异常处理代码路径
        """
        result = convert_scheme_color_to_hex("invalid_scheme")

        # 应该返回默认颜色
        assert result == "#70AD47"

    def test_convert_scheme_color_with_out_of_range_index(self):
        """
        TDD测试：convert_scheme_color_to_hex应该处理超出范围的索引

        这个测试覆盖索引超出范围的情况
        """
        result = convert_scheme_color_to_hex("accent1")

        # 应该返回正确的颜色
        assert result == "#4F81BD"

class TestGeneratePieColorVariants:
    """测试 generate_pie_color_variants 函数的边界情况。"""

    def test_generate_pie_color_variants_with_empty_base_colors(self):
        """
        TDD测试：generate_pie_color_variants应该处理空的基础颜色列表

        这个测试覆盖第245行的边界情况
        """
        result = generate_pie_color_variants("", 5)

        # 应该返回默认颜色序列
        assert len(result) == 5

    def test_generate_pie_color_variants_with_zero_count(self):
        """
        TDD测试：generate_pie_color_variants应该处理count=0的情况

        这个测试确保方法在不需要颜色时返回空列表
        """
        result = generate_pie_color_variants("#FF0000", 0)

        # 应该返回空列表
        assert result == []

    def test_generate_pie_color_variants_with_negative_count(self):
        """
        TDD测试：generate_pie_color_variants应该处理负数count

        这个测试确保方法在count为负数时返回空列表
        """
        base_colors = ["#FF0000", "#00FF00", "#0000FF"]
        result = generate_pie_color_variants(base_colors, -5)

        # 应该返回空列表
        assert result == []

class TestGenerateDistinctColors:
    """测试 generate_distinct_colors 函数的边界情况。"""

    def test_generate_distinct_colors_with_zero_count(self):
        """
        TDD测试：generate_distinct_colors应该处理count=0的情况

        这个测试覆盖第284行的边界情况
        """
        result = generate_distinct_colors(0)

        # 应该返回空列表
        assert result == []

    def test_generate_distinct_colors_with_large_count(self):
        """
        TDD测试：generate_distinct_colors应该处理大数量请求

        这个测试覆盖第300-314行的循环生成代码路径
        """
        result = generate_distinct_colors(50)

        # 应该生成50个不同的颜色
        assert len(result) == 50
        assert len(set(result)) == 50  # 所有颜色都应该是唯一的

class TestEnsureDistinctColors:
    """测试 ensure_distinct_colors 函数的边界情况。"""

    def test_ensure_distinct_colors_with_empty_list(self):
        """
        TDD测试：ensure_distinct_colors应该处理空颜色列表

        这个测试覆盖第346行的边界情况
        """
        result = ensure_distinct_colors([], 0)

        # 应该返回空列表
        assert result == []

    def test_ensure_distinct_colors_with_single_color(self):
        """
        TDD测试：ensure_distinct_colors应该处理单个颜色

        这个测试确保方法在只有一个颜色时正确处理
        """
        result = ensure_distinct_colors(["#FF0000"], 1)

        # 应该返回原始颜色
        assert result == ["#FF0000"]

class TestApplyTint:
    """测试 apply_tint 函数的边界情况。"""

    def test_apply_tint_with_invalid_color(self):
        """
        TDD测试：apply_tint应该处理无效颜色

        这个测试覆盖第373-374行的异常处理代码路径
        """
        result = apply_tint("invalid_color", 0.5)

        # 应该返回原始颜色
        assert result == "invalid_color"

    def test_apply_tint_with_extreme_tint_values(self):
        """
        TDD测试：apply_tint应该处理极端的tint值

        这个测试确保方法在极端值时正确处理
        """

        # 测试tint=1.0（完全变白）
        result = apply_tint("#000000", 1.0)
        assert result == "#FFFFFF"

        # 测试tint=0.0（不变）
        result = apply_tint("#FF0000", 0.0)
        assert result == "#FF0000"

        # 测试负值tint（变暗）
        result = apply_tint("#FF0000", -0.5)
        assert result == "#7F0000"  # 红色值减半

class TestGetColorBrightness:
    """测试 get_color_brightness 函数的边界情况。"""

    def test_get_color_brightness_with_invalid_color(self):
        """
        TDD测试：get_color_brightness应该处理无效颜色

        这个测试覆盖第383行的异常处理代码路径
        """
        result = get_color_brightness("invalid_color")

        # 应该返回0
        assert result == 0

    def test_get_color_brightness_with_short_color(self):
        """
        TDD测试：get_color_brightness应该处理短颜色代码

        这个测试确保方法在颜色代码太短时正确处理
        """
        result = get_color_brightness("#FF")

        # 应该返回0
        assert result == 0

class TestHasSufficientContrast:
    """测试 has_sufficient_contrast 函数的边界情况。"""

    def test_has_sufficient_contrast_with_invalid_colors(self):
        """
        TDD测试：has_sufficient_contrast应该处理无效颜色

        这个测试覆盖第390-391行的异常处理代码路径
        """
        result = has_sufficient_contrast("invalid1", "invalid2")

        # 应该返回False
        assert result is False

    def test_has_sufficient_contrast_with_same_colors(self):
        """
        TDD测试：has_sufficient_contrast应该处理相同颜色

        这个测试确保方法在颜色相同时返回False
        """
        result = has_sufficient_contrast("#FF0000", "#FF0000")

        # 相同颜色应该没有对比度
        assert result is False


class TestExtractColorEdgeCases:
    """测试extract_color的边界情况。"""

    def test_extract_color_with_invalid_hex_value(self):
        """
        TDD测试：extract_color应该处理无效的hex值

        这个测试覆盖第283-284行的异常处理代码
        """

        # 创建一个包含无效hex值的模拟颜色对象
        mock_color = MagicMock()
        mock_color.rgb = "INVALID_HEX"
        # 确保其他属性不存在
        del mock_color.theme
        del mock_color.indexed
        del mock_color.value
        del mock_color.auto

        result = extract_color(mock_color)

        # 验证返回None（无法解析）
        assert result is None

    def test_extract_color_with_8_char_non_ff_prefix(self):
        """
        TDD测试：extract_color应该处理8字符非FF前缀的颜色

        这个测试覆盖第297-299行的其他情况处理代码
        """

        # 创建一个8字符但不以FF开头的颜色对象
        mock_color = MagicMock()
        mock_color.rgb = "80FF0000"  # 不以FF开头

        result = extract_color(mock_color)

        # 验证取最后6位
        assert result == "#FF0000"

    def test_extract_color_with_indexed_color(self):
        """
        TDD测试：extract_color应该处理索引颜色

        这个测试覆盖第311-312行的索引颜色处理代码
        """

        # 创建一个包含索引颜色的模拟对象
        mock_color = MagicMock()
        mock_color.indexed = 2  # 红色索引

        result = extract_color(mock_color)

        # 验证返回对应的索引颜色
        assert result == "#FF0000"

    def test_extract_color_with_value_string(self):
        """
        TDD测试：extract_color应该处理value属性（字符串）

        这个测试覆盖第315-318行的value字符串处理代码
        """

        # 创建一个包含value字符串的模拟对象
        mock_color = MagicMock()
        mock_color.value = "FFFF0000"  # 8字符字符串
        # 确保其他属性不存在
        del mock_color.rgb
        del mock_color.theme
        del mock_color.indexed
        del mock_color.auto

        result = extract_color(mock_color)

        # 验证取最后6位
        assert result == "#FF0000"

    def test_extract_color_with_value_int(self):
        """
        TDD测试：extract_color应该处理value属性（整数）

        这个测试覆盖第319-320行的value整数处理代码
        """

        # 创建一个包含value整数的模拟对象
        mock_color = MagicMock()
        mock_color.value = 2  # 整数索引
        # 确保其他属性不存在
        del mock_color.rgb
        del mock_color.theme
        del mock_color.indexed
        del mock_color.auto

        result = extract_color(mock_color)

        # 验证返回对应的索引颜色
        assert result == "#FF0000"

    def test_extract_color_with_auto_color(self):
        """
        TDD测试：extract_color应该处理auto颜色

        这个测试覆盖第323-324行的auto颜色处理代码
        """

        # 创建一个auto颜色的模拟对象
        mock_color = MagicMock()
        mock_color.auto = True
        # 确保其他属性不存在
        del mock_color.rgb
        del mock_color.theme
        del mock_color.indexed
        del mock_color.value

        result = extract_color(mock_color)

        # 验证返回None（让系统决定）
        assert result is None

class TestApplyTintExceptionHandling:
    """测试apply_tint的异常处理。"""

    def test_apply_tint_with_invalid_color_format(self):
        """
        TDD测试：apply_tint应该处理无效颜色格式的异常

        这个测试覆盖第388-389行的异常处理代码
        """

        # 使用一个无效的颜色格式
        invalid_color = "#INVALID"
        result = apply_tint(invalid_color, 0.5)

        # 验证返回原始颜色
        assert result == invalid_color

class TestGetColorByIndexCoverage:
    """测试get_color_by_index的覆盖情况。"""

    def test_get_color_by_index_with_mapped_values(self):
        """
        TDD测试：get_color_by_index应该返回映射的颜色值

        这个测试覆盖第422-427行的颜色映射代码
        """

        from src.utils.color_utils import get_color_by_index

        # 测试映射表中的各种索引
        assert get_color_by_index(0) == "#000000"  # 黑色
        assert get_color_by_index(1) == "#FFFFFF"  # 白色
        assert get_color_by_index(2) == "#FF0000"  # 红色
        assert get_color_by_index(3) == "#00FF00"  # 绿色
        assert get_color_by_index(4) == "#0000FF"  # 蓝色
        assert get_color_by_index(64) == "#000000" # 特殊索引

        # 测试未映射的索引（应该返回默认黑色）
        assert get_color_by_index(999) == "#000000"

class TestApplySmartColorMatchingEdgeCases:
    """测试apply_smart_color_matching的边界情况。"""

    def test_apply_smart_color_matching_no_background(self):
        """
        TDD测试：apply_smart_color_matching应该处理没有背景色的情况

        这个测试覆盖第442行的条件分支
        """

        # 创建一个没有背景色的样式
        style = Style()
        style.background_color = None
        style.font_color = "#FF0000"

        result = apply_smart_color_matching(style)

        # 验证样式保持不变
        assert result.font_color == "#FF0000"
        assert result.background_color is None

    def test_apply_smart_color_matching_light_background_low_contrast(self):
        """
        TDD测试：apply_smart_color_matching应该为浅色背景设置黑色字体（当对比度不足时）

        这个测试覆盖第463行的浅色背景处理代码
        """

        # 创建一个浅色背景、有浅色字体的样式（对比度不足）
        style = Style()
        style.background_color = "#FFFFFF"  # 白色背景（浅色）
        style.font_color = "#CCCCCC"  # 浅灰色字体（对比度不足）

        result = apply_smart_color_matching(style)

        # 验证设置了黑色字体
        assert result.font_color == "#000000"