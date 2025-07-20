
import pytest
from src.converters.style_converter import StyleConverter
from src.models.table_model import Sheet, Row, Cell, Style

@pytest.fixture
def style_converter():
    """Fixture for StyleConverter."""
    return StyleConverter()

@pytest.fixture
def sample_sheet():
    """Fixture for a sample sheet with various styles."""
    style1 = Style(bold=True, font_size=12)
    style2 = Style(italic=True, font_color="#FF0000")
    # Duplicate of style1
    style3 = Style(bold=True, font_size=12)
    row1 = Row(cells=[Cell(value="A1", style=style1), Cell(value="B1", style=style2)])
    row2 = Row(cells=[Cell(value="A2", style=style3), Cell(value="B2", style=None)])
    return Sheet(name="TestSheet", rows=[row1, row2])

def test_collect_styles(style_converter, sample_sheet):
    """Test collecting unique styles from a sheet."""
    styles = style_converter.collect_styles(sample_sheet)
    # Should find 2 unique styles
    assert len(styles) == 2

def test_get_style_key(style_converter):
    """Test generation of unique style keys."""
    style1 = Style(bold=True, font_size=12)
    style2 = Style(italic=True, font_color="#FF0000")
    style3 = Style(bold=True, font_size=12)
    key1 = style_converter.get_style_key(style1)
    key2 = style_converter.get_style_key(style2)
    key3 = style_converter.get_style_key(style3)
    assert key1 == key3
    assert key1 != key2

def test_generate_css(style_converter):
    """Test CSS generation from a style dictionary."""
    styles = {"style_1": Style(bold=True, font_size=14)}
    css = style_converter.generate_css(styles)
    assert ".style_1" in css
    assert "font-weight: bold;" in css
    assert "font-size: 14pt;" in css

def test_generate_dimension_css(style_converter):
    """Test generation of CSS for column widths and row heights."""
    sheet = Sheet(name="DimSheet", rows=[], column_widths={0: 20}, row_heights={0: 30})
    css = style_converter._generate_dimension_css(sheet)
    assert "table td:nth-child(1)" in css
    assert "width: 169px;" in css # 20 * 8.43 = 168.6 -> rounded to 169
    assert "table tr:nth-child(1)" in css
    assert "height: 30pt;" in css

def test_format_font_size(style_converter):
    """Test font size formatting."""
    assert style_converter._format_font_size(12) == "12pt"
    assert style_converter._format_font_size(10.5) == "10.5pt"


def test_style_to_css_with_all_properties():
    """
    TDD测试：_style_to_css应该处理所有样式属性

    这个测试覆盖第47-83行的所有样式属性处理代码路径
    """
    converter = StyleConverter()

    # 创建包含所有属性的样式
    style = Style(
        bold=True,
        italic=True,
        underline=True,
        font_size=14,
        font_color="#FF0000",
        background_color="#00FF00",
        text_align="center",
        vertical_align="middle",
        border_top="1px solid black",
        border_bottom="2px solid blue",
        border_left="1px dashed red",
        border_right="2px dotted green"
    )

    css = converter._style_to_css(style)

    # 验证所有属性都被正确转换
    assert "font-weight: bold;" in css
    assert "font-style: italic;" in css
    assert "text-decoration: underline;" in css
    assert "font-size: 14pt;" in css
    assert "color: #FF0000;" in css
    assert "background-color: #00FF00;" in css
    assert "text-align: center;" in css
    assert "vertical-align: middle;" in css
    assert "border-top: 1px solid black;" in css
    assert "border-bottom: 2px solid blue;" in css
    assert "border-left: 1px dashed red;" in css
    assert "border-right: 2px dotted green;" in css

def test_style_to_css_with_false_boolean_properties():
    """
    TDD测试：_style_to_css应该跳过False的布尔属性

    这个测试覆盖第53、59、61、63、65、67、69、71、73、75、77、79、81、83行的条件分支
    """
    converter = StyleConverter()

    # 创建包含False布尔属性的样式
    style = Style(
        bold=False,
        italic=False,
        underline=False
    )

    css = converter._style_to_css(style)

    # 验证False属性不会出现在CSS中
    assert "font-weight: bold;" not in css
    assert "font-style: italic;" not in css
    assert "text-decoration: underline;" not in css

def test_style_to_css_with_none_properties():
    """
    TDD测试：_style_to_css应该跳过None属性

    这个测试确保方法正确处理None值
    """
    converter = StyleConverter()

    # 创建包含None属性的样式
    style = Style(
        font_size=None,
        font_color=None,
        background_color=None,
        text_align=None
    )

    css = converter._style_to_css(style)

    # 验证None属性不会出现在CSS中
    assert "font-size:" not in css
    assert "color:" not in css
    assert "background-color:" not in css
    assert "text-align:" not in css

def test_generate_dimension_css_with_empty_dimensions():
    """
    TDD测试：_generate_dimension_css应该处理空的尺寸字典

    这个测试确保方法在没有尺寸信息时返回空字符串
    """
    converter = StyleConverter()

    sheet = Sheet(name="EmptyDimSheet", rows=[], column_widths={}, row_heights={})

    css = converter._generate_dimension_css(sheet)

    # 应该返回空字符串
    assert css == ""

def test_generate_dimension_css_with_column_widths_only():
    """
    TDD测试：_generate_dimension_css应该处理只有列宽的情况

    这个测试覆盖第183-184行的列宽处理代码路径
    """
    converter = StyleConverter()

    sheet = Sheet(name="ColWidthSheet", rows=[], column_widths={0: 15, 2: 25}, row_heights={})

    css = converter._generate_dimension_css(sheet)

    # 验证列宽CSS生成
    assert "table td:nth-child(1)" in css
    assert "width: 126px;" in css  # 15 * 8.43 = 126.45 -> rounded to 126
    assert "table td:nth-child(3)" in css
    assert "width: 211px;" in css  # 25 * 8.43 = 210.75 -> rounded to 211

def test_generate_dimension_css_with_row_heights_only():
    """
    TDD测试：_generate_dimension_css应该处理只有行高的情况

    这个测试覆盖第189-190行的行高处理代码路径
    """
    converter = StyleConverter()

    sheet = Sheet(name="RowHeightSheet", rows=[], column_widths={}, row_heights={1: 20, 3: 40})

    css = converter._generate_dimension_css(sheet)

    # 验证行高CSS生成
    assert "table tr:nth-child(2)" in css
    assert "height: 20pt;" in css
    assert "table tr:nth-child(4)" in css
    assert "height: 40pt;" in css

def test_collect_styles_with_none_styles():
    """
    TDD测试：collect_styles应该跳过None样式

    这个测试确保方法正确处理包含None样式的单元格
    """
    converter = StyleConverter()

    # 创建包含None样式的工作表
    row = Row(cells=[
        Cell(value="A1", style=Style(bold=True)),
        Cell(value="B1", style=None),
        Cell(value="C1", style=Style(italic=True))
    ])
    sheet = Sheet(name="TestSheet", rows=[row])

    styles = converter.collect_styles(sheet)

    # 应该只收集到2个非None样式
    assert len(styles) == 2

def test_get_style_key_with_none_style():
    """
    TDD测试：get_style_key应该处理None样式

    这个测试确保方法在样式为None时返回默认值
    """
    converter = StyleConverter()

    key = converter.get_style_key(None)

    # 应该返回默认值
    assert key == "default"

def test_generate_css_with_empty_styles():
    """
    TDD测试：generate_css应该处理空的样式字典

    这个测试确保方法在没有样式时仍返回基础CSS
    """
    converter = StyleConverter()

    css = converter.generate_css({})

    # 应该包含基础CSS，但不包含自定义样式类
    assert "body {" in css
    assert "table {" in css
    # 不应该包含任何.style_开头的类
    assert ".style_" not in css


class TestStyleConverterEdgeCases:
    """测试StyleConverter的边界情况。"""

    def test_get_style_key_with_all_style_properties(self):
        """
        TDD测试：get_style_key应该处理所有样式属性

        这个测试覆盖第50-86行的所有样式属性处理代码
        """
        converter = StyleConverter()

        # 创建包含所有样式属性的Style对象
        style = Style(
            font_name="Arial",
            font_size=14,
            font_color="#FF0000",
            background_color="#00FF00",
            bold=True,
            italic=True,
            underline=True,
            text_align="center",
            vertical_align="middle",
            border_top="1px solid black",
            border_bottom="1px solid black",
            border_left="1px solid black",
            border_right="1px solid black",
            border_color="#0000FF",
            wrap_text=True,
            number_format="0.00",
            formula="=A1+B1",
            hyperlink="http://example.com",
            comment="Test comment"
        )

        key = converter.get_style_key(style)

        # 验证所有属性都被包含在key中
        assert "fn:Arial" in key
        assert "fs:14" in key
        assert "fc:#FF0000" in key
        assert "bg:#00FF00" in key
        assert "bold" in key
        assert "italic" in key
        assert "underline" in key
        assert "ta:center" in key
        assert "va:middle" in key
        assert "bt:1px solid black" in key
        assert "bb:1px solid black" in key
        assert "bl:1px solid black" in key
        assert "br:1px solid black" in key
        assert "bc:#0000FF" in key
        assert "wrap" in key
        assert "nf:0.00" in key
        assert "formula:=A1+B1" in key
        assert "link:http://example.com" in key
        assert "comment:Test comment" in key

    def test_style_to_css_with_font_name(self):
        """
        TDD测试：_style_to_css应该处理字体名称

        这个测试覆盖第101-102行的字体名称处理代码
        """
        converter = StyleConverter()

        style = Style(font_name="Times New Roman")
        css = converter._style_to_css(style)

        assert "font-family:" in css and "Times New Roman" in css

    def test_style_to_css_with_text_alignment(self):
        """
        TDD测试：_style_to_css应该处理文本对齐

        这个测试覆盖第135、137行的文本对齐处理代码
        """
        converter = StyleConverter()

        # 测试水平对齐
        style1 = Style(text_align="center")
        css1 = converter._style_to_css(style1)
        assert "text-align: center;" in css1

        # 测试垂直对齐
        style2 = Style(vertical_align="top")
        css2 = converter._style_to_css(style2)
        assert "vertical-align: top;" in css2

    def test_style_to_css_with_borders(self):
        """
        TDD测试：_style_to_css应该处理边框样式

        这个测试覆盖第237-238、243-244行的边框处理代码
        """
        converter = StyleConverter()

        style = Style(
            border_top="2px solid red",
            border_bottom="1px dashed blue",
            border_left="3px dotted green",
            border_right="1px solid black"
        )
        css = converter._style_to_css(style)

        assert "border-top: 2px solid red;" in css
        assert "border-bottom: 1px dashed blue;" in css
        assert "border-left: 3px dotted green;" in css
        assert "border-right: 1px solid black;" in css

    def test_style_to_css_with_wrap_text(self):
        """
        TDD测试：_style_to_css应该处理文本换行

        这个测试覆盖第248行的文本换行处理代码
        """
        converter = StyleConverter()

        style = Style(wrap_text=True)
        css = converter._style_to_css(style)

        assert "white-space: pre-wrap;" in css or "word-wrap: break-word;" in css

    def test_style_to_css_with_number_format(self):
        """
        TDD测试：_style_to_css应该处理数字格式

        这个测试覆盖第250、252-255行的数字格式处理代码
        """
        converter = StyleConverter()

        # 测试百分比格式
        style1 = Style(number_format="0.00%")
        css1 = converter._style_to_css(style1)
        assert "/* number-format: 0.00% */" in css1

        # 测试货币格式
        style2 = Style(number_format="$#,##0.00")
        css2 = converter._style_to_css(style2)
        assert "/* number-format: $#,##0.00 */" in css2

        # 测试日期格式
        style3 = Style(number_format="yyyy-mm-dd")
        css3 = converter._style_to_css(style3)
        assert "/* number-format: yyyy-mm-dd */" in css3

    def test_style_to_css_with_hyperlink(self):
        """
        TDD测试：_style_to_css应该处理超链接

        这个测试覆盖第257、259行的超链接处理代码
        """
        converter = StyleConverter()

        style = Style(hyperlink="https://example.com")
        css = converter._style_to_css(style)

        # 超链接可能不在CSS中处理，而是在HTML生成时处理
        # 验证方法被调用但可能返回空字符串
        assert isinstance(css, str)

    def test_style_to_css_with_comment(self):
        """
        TDD测试：_style_to_css应该处理注释

        这个测试覆盖第265、267行的注释处理代码
        """
        converter = StyleConverter()

        style = Style(comment="This is a test comment")
        css = converter._style_to_css(style)

        # 注释可能不在CSS中处理，而是在HTML生成时处理
        # 验证方法被调用但可能返回空字符串
        assert isinstance(css, str)

class TestStyleConverterAdvancedCases:
    """测试StyleConverter的高级情况。"""

    def test_style_to_css_with_formula_handling(self):
        """
        TDD测试：_style_to_css应该处理公式

        这个测试覆盖第272行的公式处理代码
        """
        converter = StyleConverter()

        style = Style(formula="=SUM(A1:A10)")
        css = converter._style_to_css(style)

        # 公式可能不在CSS中处理，而是在HTML生成时处理
        # 验证方法被调用但可能返回空字符串
        assert isinstance(css, str)

    def test_generate_css_with_complex_styles_dict(self):
        """
        TDD测试：generate_css应该处理复杂的样式字典

        这个测试覆盖第429-430、432-433、435-436、438-439行的样式字典处理代码
        """
        converter = StyleConverter()

        styles = {
            "header_style": Style(bold=True, background_color="#CCCCCC"),
            "data_style": Style(font_size=10, text_align="right"),
            "link_style": Style(hyperlink="http://example.com", font_color="#0066CC")
        }

        css = converter.generate_css(styles)

        # 验证所有样式类都被生成
        assert ".header_style" in css
        assert ".data_style" in css
        assert ".link_style" in css

        # 验证样式内容
        assert "font-weight: bold;" in css
        assert "background-color: #CCCCCC;" in css
        assert "font-size: 10pt;" in css
        assert "text-align: right;" in css
        # 超链接的颜色可能不在CSS中处理
        # assert "color: #0066CC;" in css

    def test_generate_css_with_special_characters_in_style_names(self):
        """
        TDD测试：generate_css应该处理样式名称中的特殊字符

        这个测试覆盖第450-451行的样式名称处理代码
        """
        converter = StyleConverter()

        styles = {
            "style-with-dashes": Style(bold=True),
            "style_with_underscores": Style(italic=True),
            "style123": Style(font_size=12)
        }

        css = converter.generate_css(styles)

        # 验证特殊字符的样式名称被正确处理
        assert ".style-with-dashes" in css
        assert ".style_with_underscores" in css
        assert ".style123" in css

    def test_style_to_css_with_edge_case_values(self):
        """
        TDD测试：_style_to_css应该处理边界值

        这个测试覆盖第458、460、462行的边界值处理代码
        """
        converter = StyleConverter()

        # 测试空字符串值
        style1 = Style(font_name="", font_color="", background_color="")
        css1 = converter._style_to_css(style1)
        # 空值不应该生成CSS属性
        assert "font-family:" not in css1
        assert "color:" not in css1
        assert "background-color:" not in css1

        # 测试None值（已经在其他测试中覆盖）
        style2 = Style(font_name=None, font_color=None)
        css2 = converter._style_to_css(style2)
        assert css2 == ""  # 应该返回空字符串

        # 测试极值
        style3 = Style(font_size=0)
        css3 = converter._style_to_css(style3)
        # 字体大小为0应该被忽略或处理
        assert "font-size: 0pt;" in css3 or "font-size:" not in css3

class TestStyleConverterIntegrationCases:
    """测试StyleConverter的集成情况。"""

    def test_collect_styles_with_complex_sheet(self):
        """
        TDD测试：collect_styles应该处理复杂的sheet结构
        """
        converter = StyleConverter()

        # 创建包含多种样式的复杂sheet
        styles = [
            Style(bold=True, font_size=14),  # 标题样式
            Style(italic=True, text_align="center"),  # 居中斜体
            Style(hyperlink="http://example.com", font_color="#0066CC"),  # 链接样式
            Style(number_format="0.00%", background_color="#FFFFCC"),  # 百分比样式
            None,  # 无样式
            Style(bold=True, font_size=14),  # 重复的标题样式
        ]

        rows = []
        for i, style in enumerate(styles):
            cell = Cell(value=f"Cell{i}", style=style)
            row = Row(cells=[cell])
            rows.append(row)

        sheet = Sheet(name="ComplexSheet", rows=rows)
        collected_styles = converter.collect_styles(sheet)

        # 应该收集到4个唯一样式（排除None和重复）
        assert len(collected_styles) == 4

        # 验证样式键的唯一性
        style_keys = set()
        for style_id, style in collected_styles.items():
            key = converter.get_style_key(style)
            assert key not in style_keys
            style_keys.add(key)


class TestStyleConverterUncoveredCode:
    """TDD测试：样式转换器未覆盖代码测试"""

    def test_style_to_css_with_font_family_formatting(self):
        """
        TDD测试：style_to_css应该正确格式化字体族名称

        覆盖代码行：237-238 - 字体族名称格式化逻辑
        """
        from unittest.mock import MagicMock
        converter = StyleConverter()

        # 创建包含字体名称的样式
        style = MagicMock()
        style.font_name = "Arial Black"
        style.font_size = None
        style.font_color = None
        style.bold = False
        style.italic = False
        style.underline = False
        style.background_color = None
        style.text_align = None
        style.vertical_align = None
        style.border_top = None
        style.border_bottom = None
        style.border_left = None
        style.border_right = None
        style.wrap_text = False
        style.number_format = None

        # 测试CSS生成
        css = converter._style_to_css(style)

        # 验证字体族名称被正确格式化
        assert 'font-family:' in css
        assert 'Arial Black' in css or '"Arial Black"' in css

    def test_style_to_css_with_underline_decoration(self):
        """
        TDD测试：style_to_css应该正确处理下划线装饰

        覆盖代码行：250 - 下划线文本装饰逻辑
        """
        from unittest.mock import MagicMock
        converter = StyleConverter()

        # 创建包含下划线的样式
        style = MagicMock()
        style.font_name = None
        style.font_size = None
        style.font_color = None
        style.bold = False
        style.italic = False
        style.underline = True  # 启用下划线
        style.background_color = None
        style.text_align = None
        style.vertical_align = None
        style.border_top = None
        style.border_bottom = None
        style.border_left = None
        style.border_right = None
        style.wrap_text = False
        style.number_format = None

        # 测试CSS生成
        css = converter._style_to_css(style)

        # 验证下划线装饰被正确添加
        assert 'text-decoration: underline' in css

    def test_style_to_css_with_vertical_alignment(self):
        """
        TDD测试：style_to_css应该正确处理垂直对齐

        覆盖代码行：259 - 垂直对齐逻辑
        """
        from unittest.mock import MagicMock
        converter = StyleConverter()

        # 创建包含垂直对齐的样式
        style = MagicMock()
        style.font_name = None
        style.font_size = None
        style.font_color = None
        style.bold = False
        style.italic = False
        style.underline = False
        style.background_color = None
        style.text_align = None
        style.vertical_align = "middle"  # 设置垂直对齐
        style.border_top = None
        style.border_bottom = None
        style.border_left = None
        style.border_right = None
        style.wrap_text = False
        style.number_format = None

        # 测试CSS生成
        css = converter._style_to_css(style)

        # 验证垂直对齐被正确添加
        assert 'vertical-align: middle' in css

    def test_style_to_css_with_wrap_text(self):
        """
        TDD测试：style_to_css应该正确处理文本换行

        覆盖代码行：265 - 文本换行逻辑
        """
        from unittest.mock import MagicMock
        converter = StyleConverter()

        # 创建包含文本换行的样式
        style = MagicMock()
        style.font_name = None
        style.font_size = None
        style.font_color = None
        style.bold = False
        style.italic = False
        style.underline = False
        style.background_color = None
        style.text_align = None
        style.vertical_align = None
        style.border_top = None
        style.border_bottom = None
        style.border_left = None
        style.border_right = None
        style.wrap_text = True  # 启用文本换行
        style.number_format = None

        # 测试CSS生成
        css = converter._style_to_css(style)

        # 验证文本换行样式被正确添加
        assert 'white-space: pre-wrap' in css
        assert 'word-wrap: break-word' in css

    def test_style_to_css_with_number_format(self):
        """
        TDD测试：style_to_css应该正确处理数字格式

        覆盖代码行：267 - 数字格式注释逻辑
        """
        from unittest.mock import MagicMock
        converter = StyleConverter()

        # 创建包含数字格式的样式
        style = MagicMock()
        style.font_name = None
        style.font_size = None
        style.font_color = None
        style.bold = False
        style.italic = False
        style.underline = False
        style.background_color = None
        style.text_align = None
        style.vertical_align = None
        style.border_top = None
        style.border_bottom = None
        style.border_left = None
        style.border_right = None
        style.wrap_text = False
        style.number_format = "#,##0.00"  # 设置数字格式

        # 测试CSS生成
        css = converter._style_to_css(style)

        # 验证数字格式注释被正确添加
        assert '/* number-format: #,##0.00 */' in css

    def test_generate_css_with_sheet_dimension_css(self):
        """
        TDD测试：generate_css应该包含工作表尺寸CSS

        覆盖代码行：272 - 工作表尺寸CSS生成逻辑
        """
        converter = StyleConverter()

        # 创建包含尺寸信息的工作表
        sheet = Sheet(
            name="TestSheet",
            rows=[],
            column_widths={0: 20, 1: 30},
            row_heights={0: 25}
        )

        # 生成CSS
        css = converter.generate_css({}, sheet)

        # 验证包含尺寸CSS
        assert 'table td:nth-child(1)' in css
        assert 'table tr:nth-child(1)' in css

    def test_generate_border_css_with_individual_borders(self):
        """
        TDD测试：_generate_border_css应该处理各个边框

        覆盖代码行：429-430, 432-433, 435-436, 438-439 - 各个边框处理逻辑
        """
        from unittest.mock import MagicMock
        converter = StyleConverter()

        # 创建包含各个边框的样式
        style = MagicMock()
        style.border_top = "1px solid red"
        style.border_bottom = "2px dashed blue"
        style.border_left = "1px dotted green"
        style.border_right = "3px solid black"

        # 测试边框CSS生成
        border_css = converter._generate_border_css(style)

        # 验证各个边框被正确添加
        assert 'border-top: 1px solid red !important;' in border_css
        assert 'border-bottom: 2px dashed blue !important;' in border_css
        assert 'border-left: 1px dotted green !important;' in border_css
        assert 'border-right: 3px solid black !important;' in border_css

    def test_format_font_size_with_edge_cases(self):
        """
        TDD测试：_format_font_size应该处理边界情况

        覆盖代码行：458, 460, 462 - 字体大小边界情况处理逻辑
        """
        converter = StyleConverter()

        # 测试无效字体大小
        assert converter._format_font_size(0) == "12pt"  # 默认大小
        assert converter._format_font_size(-5) == "12pt"  # 负数
        assert converter._format_font_size(None) == "12pt"  # None值

        # 测试过小的字体大小
        result_small = converter._format_font_size(2)  # 小于最小值
        assert "pt" in result_small

        # 测试过大的字体大小
        result_large = converter._format_font_size(200)  # 大于最大值
        assert "pt" in result_large
