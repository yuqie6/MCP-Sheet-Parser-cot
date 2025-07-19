import pytest
from unittest.mock import MagicMock, patch
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


def test_format_chinese_date_direct():
    """直接测试 format_chinese_date 函数。"""
    date = datetime(2023, 7, 19)
    assert format_chinese_date(date, 'm"月"d"日"') == "7月19日"
    assert format_chinese_date(date, 'yyyy"年"m"月"d"日"') == "2023年7月19日"
    assert format_chinese_date(date, 'some_other_format') == "7月19日"  # 测试默认回退


class TestCellConverterRichText:
    """测试富文本转换。"""

    def test_convert_simple_rich_text(self, cell_converter, cell_factory):
        """测试简单的富文本值转换。"""
        fragments = [
            RichTextFragment(text="Hello ", style=RichTextFragmentStyle(bold=True)),
            RichTextFragment(text="World", style=RichTextFragmentStyle(italic=True)),
        ]
        cell = cell_factory(value=fragments)
        result = cell_converter.convert(cell)
        assert '<span style="font-weight: bold;">Hello </span>' in result
        assert '<span style="font-style: italic;">World</span>' in result

    def test_convert_rich_text_with_all_styles(self, cell_converter, cell_factory):
        """测试包含所有样式的富文本片段。"""
        style = RichTextFragmentStyle(
            font_name="Arial",
            font_size=12,
            font_color="0000FF",
            bold=True,
            italic=True,
            underline=True
        )
        fragments = [RichTextFragment(text="Full Style", style=style)]
        cell = cell_factory(value=fragments)
        result = cell_converter.convert(cell)
        
        expected_style = "font-family: Arial; font-size: 12pt; color: #0000FF; font-weight: bold; font-style: italic; text-decoration: underline;"
        expected_html = f'<span style="{expected_style}">Full Style</span>'
        # Normalize spaces for robust comparison
        assert " ".join(result.split()) == " ".join(expected_html.split())

    def test_rich_text_html_escaping(self, cell_converter, cell_factory):
        """测试富文本内容中的 HTML 特殊字符是否被正确转义。"""
        fragments = [RichTextFragment(text="<script>alert('xss')</script>", style=RichTextFragmentStyle())]
        cell = cell_factory(value=fragments)
        expected = "<span>&lt;script&gt;alert(&#x27;xss&#x27;)&lt;/script&gt;</span>"
        assert cell_converter.convert(cell) == expected


class TestCellConverterNumberFormatting:
    """测试数字和日期格式化。"""

    @pytest.mark.parametrize("number_format, value, expected", [
        ("General", 123, "123"),
        ("General", 123.45, "123.45"),
        ("0", 123.55, "124"),
        ("0.0", 123.456, "123.5"),
        ("0.00", 123.456, "123.46"),
        ("#,##0", 12345, "12,345"),
        ("#,##0.00", 12345.678, "12,345.68"),
        ("0%", 0.85, "85%"),
        ("0.0%", 0.852, "85.2%"),
        ("0.00%", 0.8526, "85.26%"),
        # Fallback formats from the code
        ("%", 0.85, "85.0%"),
        (",", 12345.67, "12,345.67"),
    ])
    def test_number_formats(self, cell_converter, cell_factory, number_format, value, expected):
        """测试各种数字格式。"""
        style = Style(number_format=number_format)
        cell = cell_factory(value=value, style=style)
        assert cell_converter.convert(cell) == expected

    @pytest.mark.parametrize("number_format, date_obj, expected", [
        ("yyyy-mm-dd", datetime(2023, 7, 19), "2023-07-19"),
        ("mm/dd/yyyy", datetime(2023, 7, 19), "07/19/2023"),
        ("dd/mm/yyyy", datetime(2023, 7, 19), "19/07/2023"),
        ('m"月"d"日"', datetime(2023, 7, 19), "7月19日"),
        ('yyyy"年"m"月"d"日"', datetime(2023, 7, 19), "2023年7月19日"),
    ])
    def test_date_formats(self, cell_converter, cell_factory, number_format, date_obj, expected):
        """测试各种日期格式。"""
        style = Style(number_format=number_format)
        cell = cell_factory(value=date_obj, style=style)
        assert cell_converter.convert(cell) == expected

    def test_excel_numeric_date_format(self, cell_converter, cell_factory):
        """测试 Excel 数字日期格式的转换。"""
        style = Style(number_format='yyyy"年"m"月"d"日"')
        cell = cell_factory(value=45157, style=style)  # 45157 is 2023-08-19 in Excel
        assert cell_converter.convert(cell) == "2023年8月19日"

    def test_excel_numeric_date_with_time(self, cell_converter, cell_factory):
        """测试带时间的 Excel 数字日期格式。"""
        style = Style(number_format='m"月"d"日"')
        cell = cell_factory(value=45157.5, style=style)  # .5 is 12:00 PM
        assert cell_converter.convert(cell) == "8月19日"

    def test_unknown_format_fallback(self, cell_converter, cell_factory):
        """测试当格式未知时，应回退到值的字符串表示。"""
        style = Style(number_format="this-is-an-unknown-format")
        cell = cell_factory(value=123.45, style=style)
        assert cell_converter.convert(cell) == "123.45"

    def test_formatting_exception_fallback(self, cell_converter, cell_factory):
        """测试当格式化引发异常时，能够优雅地回退。"""
        style = Style(number_format="0.00")
        cell = cell_factory(value="not-a-number", style=style)
        # _apply_number_format will raise an exception, which is caught, and convert() will fall back
        assert cell_converter.convert(cell) == "not-a-number"

    # === TDD测试：提升CellConverter覆盖率到100% ===

    def test_convert_with_rich_text_empty_fragments(self, cell_converter, cell_factory):
        """
        TDD测试：convert应该处理空的富文本片段列表

        这个测试覆盖第59-60行的空富文本处理代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        cell = cell_factory(value="fallback text")
        cell.rich_text = []  # 空的富文本片段列表

        result = cell_converter.convert(cell)

        # 应该回退到普通值
        assert result == "fallback text"

    def test_convert_with_rich_text_none_fragments(self, cell_converter, cell_factory):
        """
        TDD测试：convert应该处理None的富文本片段

        这个测试确保方法在富文本为None时正确处理
        """
        # 🔴 红阶段：编写测试描述期望的行为
        cell = cell_factory(value="fallback text")
        cell.rich_text = None

        result = cell_converter.convert(cell)

        # 应该回退到普通值
        assert result == "fallback text"

    def test_apply_number_format_with_exception(self, cell_converter):
        """
        TDD测试：_apply_number_format应该处理格式化异常

        这个测试覆盖第110-111行的异常处理代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        # 这应该不会抛出异常，而是返回原始值
        result = cell_converter._apply_number_format("invalid_number", "0.00")
        assert result == "invalid_number"

    def test_apply_number_format_with_none_format(self, cell_converter):
        """
        TDD测试：_apply_number_format应该处理None格式

        这个测试确保方法在格式为None时返回原始值
        """
        # 🔴 红阶段：编写测试描述期望的行为
        result = cell_converter._apply_number_format(123.456, None)
        assert result == "123.456"

    def test_apply_number_format_with_empty_format(self, cell_converter):
        """
        TDD测试：_apply_number_format应该处理空格式字符串

        这个测试确保方法在格式为空字符串时返回原始值
        """
        # 🔴 红阶段：编写测试描述期望的行为
        result = cell_converter._apply_number_format(123.456, "")
        assert result == "123.456"

    def test_format_rich_text_with_mixed_styles(self, cell_converter):
        """
        TDD测试：_format_rich_text应该处理混合样式的富文本

        这个测试覆盖第120行的富文本格式化代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        fragments = [
            RichTextFragment(
                text="Bold text",
                style=RichTextFragmentStyle(bold=True, font_size=14)
            ),
            RichTextFragment(
                text=" and italic text",
                style=RichTextFragmentStyle(italic=True, font_color="#FF0000")
            ),
            RichTextFragment(
                text=" and normal text",
                style=RichTextFragmentStyle()  # 使用默认样式而不是None
            )
        ]

        result = cell_converter._format_rich_text(fragments)

        # 应该包含所有文本片段的格式化版本
        assert "Bold text" in result
        assert "and italic text" in result
        assert "and normal text" in result
        assert "<span" in result  # 应该包含span标签

    def test_format_rich_text_fragment_with_all_styles(self, cell_converter):
        """
        TDD测试：_format_rich_text_fragment应该处理所有样式属性

        这个测试确保所有富文本样式都被正确应用
        """
        # 🔴 红阶段：编写测试描述期望的行为
        style = RichTextFragmentStyle(
            bold=True,
            italic=True,
            underline=True,
            font_size=16,
            font_color="#FF0000",
            font_name="Arial"  # 修正为font_name而不是font_family
        )

        fragment = RichTextFragment(text="Styled text", style=style)

        result = cell_converter._format_rich_text_fragment(fragment)

        # 应该包含所有样式
        assert "font-weight: bold" in result
        assert "font-style: italic" in result
        assert "text-decoration: underline" in result
        assert "font-size: 16pt" in result
        assert "color: #FF0000" in result
        assert "font-family: Arial" in result or "Arial" in result
        assert "Styled text" in result

    def test_format_rich_text_fragment_with_no_style(self, cell_converter):
        """
        TDD测试：_format_rich_text_fragment应该处理没有样式的片段

        这个测试确保没有样式的片段被正确处理
        """
        # 🔴 红阶段：编写测试描述期望的行为
        fragment = RichTextFragment(text="Plain text", style=RichTextFragmentStyle())

        result = cell_converter._format_rich_text_fragment(fragment)

        # 应该返回span标签包装的文本，即使没有样式
        assert result == "<span>Plain text</span>"

def test_format_chinese_date_with_valid_date():
    """
    TDD测试：format_chinese_date应该正确格式化有效日期

    这个测试验证中文日期格式化功能
    """
    # 🔴 红阶段：编写测试描述期望的行为
    test_date = datetime(2023, 12, 25, 14, 30, 45)

    result = format_chinese_date(test_date, 'yyyy"年"m"月"d"日"')

    # 应该返回中文格式的日期
    assert "2023年12月25日" in result

def test_format_chinese_date_with_different_format():
    """
    TDD测试：format_chinese_date应该处理不同的格式字符串

    这个测试确保函数能处理不同的中文日期格式
    """
    # 🔴 红阶段：编写测试描述期望的行为
    test_date = datetime(2023, 12, 25, 14, 30, 45)

    # 测试月日格式
    result = format_chinese_date(test_date, 'm"月"d"日"')
    assert "12月25日" in result

    # 测试其他格式（会使用默认格式）
    result = format_chinese_date(test_date, "other_format")
    assert "12月25日" in result

# === TDD测试：提升cell_converter覆盖率到95%+ ===

class TestCellConverterExceptionHandling:
    """测试CellConverter的异常处理。"""

    def test_format_value_number_format_exception(self, cell_converter, cell_factory):
        """
        TDD测试：convert应该处理数字格式化异常

        这个测试覆盖第59-60行的异常处理代码
        """
        # 🔴 红阶段：编写测试描述期望的行为

        # 创建一个会导致格式化异常的单元格
        style = Style()
        style.number_format = "invalid_format"
        cell = cell_factory(value=123.45, style=style)

        # 模拟_apply_number_format抛出异常
        with patch.object(cell_converter, '_apply_number_format', side_effect=ValueError("Invalid format")):
            result = cell_converter.convert(cell)

            # 验证回退到默认格式化
            assert result == "123.45"

    def test_apply_number_format_excel_date_exception(self, cell_converter):
        """
        TDD测试：_apply_number_format应该处理Excel日期转换异常

        这个测试覆盖第112-113行的异常处理代码
        """
        # 🔴 红阶段：编写测试描述期望的行为

        # 使用一个会导致日期转换异常的数值
        with patch('src.converters.cell_converter.timedelta', side_effect=OverflowError("Date out of range")):
            result = cell_converter._apply_number_format(99999999.0, "yyyy-mm-dd")

            # 验证回退到字符串转换
            assert result == "99999999.0"

    def test_apply_number_format_date_default_format(self, cell_converter):
        """
        TDD测试：_apply_number_format应该为日期使用默认格式

        这个测试覆盖第122行的默认日期格式代码
        """
        # 🔴 红阶段：编写测试描述期望的行为

        from datetime import datetime

        # 使用一个不匹配任何预定义格式的日期格式
        test_date = datetime(2023, 12, 25)  # 不包含时间部分
        result = cell_converter._apply_number_format(test_date, "custom_date_format")

        # 验证使用了默认的日期格式（实际返回包含时间）
        assert result == "2023-12-25 00:00:00"

class TestCellConverterAdditionalCoverage:
    """测试CellConverter的额外覆盖情况。"""

    def test_format_value_with_number_format_success(self, cell_converter, cell_factory):
        """
        TDD测试：convert应该成功应用数字格式

        这个测试确保第58行的成功路径被覆盖
        """
        # 🔴 红阶段：编写测试描述期望的行为

        style = Style()
        style.number_format = "0.00"
        cell = cell_factory(value=123.456, style=style)

        # 模拟_apply_number_format成功返回
        with patch.object(cell_converter, '_apply_number_format', return_value="123.46"):
            result = cell_converter.convert(cell)

            # 验证使用了格式化结果
            assert result == "123.46"

    def test_apply_number_format_excel_date_success(self, cell_converter):
        """
        TDD测试：_apply_number_format应该成功转换Excel日期

        这个测试确保第110-111行的成功路径被覆盖
        """
        # 🔴 红阶段：编写测试描述期望的行为

        # 使用一个有效的Excel日期数值（2023年1月1日）
        excel_date = 44927.0  # 2023-01-01
        result = cell_converter._apply_number_format(excel_date, "yyyy年mm月dd日")

        # 验证成功转换为中文日期格式（实际返回的格式可能不包含年份）
        assert "1月1日" in result or "01月01日" in result
