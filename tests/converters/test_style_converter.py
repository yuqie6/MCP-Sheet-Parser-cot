
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

# === TDD测试：提升StyleConverter覆盖率 ===

def test_style_to_css_with_all_properties():
    """
    TDD测试：_style_to_css应该处理所有样式属性

    这个测试覆盖第47-83行的所有样式属性处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
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
    # 🔴 红阶段：编写测试描述期望的行为
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
    # 🔴 红阶段：编写测试描述期望的行为
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
    # 🔴 红阶段：编写测试描述期望的行为
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
    # 🔴 红阶段：编写测试描述期望的行为
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
    # 🔴 红阶段：编写测试描述期望的行为
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
    # 🔴 红阶段：编写测试描述期望的行为
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
    # 🔴 红阶段：编写测试描述期望的行为
    converter = StyleConverter()

    key = converter.get_style_key(None)

    # 应该返回默认值
    assert key == "default"

def test_generate_css_with_empty_styles():
    """
    TDD测试：generate_css应该处理空的样式字典

    这个测试确保方法在没有样式时仍返回基础CSS
    """
    # 🔴 红阶段：编写测试描述期望的行为
    converter = StyleConverter()

    css = converter.generate_css({})

    # 应该包含基础CSS，但不包含自定义样式类
    assert "body {" in css
    assert "table {" in css
    # 不应该包含任何.style_开头的类
    assert ".style_" not in css
