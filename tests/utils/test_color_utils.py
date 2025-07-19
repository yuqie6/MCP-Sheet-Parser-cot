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