

import pytest
from src.utils.html_utils import (
    escape_html,
    generate_style_attribute,
    generate_class_attribute,
    create_html_element,
    create_table_cell,
    create_svg_element,
    compact_html
)

def test_escape_html():
    """Test HTML escaping."""
    assert escape_html("<script>alert('xss')</script>") == "&lt;script&gt;alert(&#x27;xss&#x27;)&lt;/script&gt;"

def test_generate_style_attribute():
    """Test style attribute generation."""
    assert generate_style_attribute(["color: red;", "font-size: 12px;"]) == 'style="color: red; font-size: 12px;"'

def test_generate_class_attribute():
    """Test class attribute generation."""
    assert generate_class_attribute(["class1", "class2"]) == 'class="class1 class2"'

def test_create_html_element():
    """Test HTML element creation."""
    html = create_html_element("div", "Hello", attributes={"id": "my-div"}, css_classes=["container"], inline_styles={"color": "blue"})
    assert 'id="my-div"' in html
    assert 'class="container"' in html
    assert 'style="color: blue;"' in html
    assert ">Hello</div>" in html

def test_create_table_cell():
    """Test table cell creation."""
    th = create_table_cell("Header", is_header=True, colspan=2)
    assert th.startswith("<th")
    assert 'colspan="2"' in th
    td = create_table_cell("Data")
    assert td.startswith("<td")

def test_create_svg_element():
    """Test SVG element creation."""
    svg = create_svg_element(100, 50, "<circle />")
    assert 'width="100px"' in svg
    assert 'height="50px"' in svg
    assert '<circle />' in svg

def test_compact_html():
    """Test HTML compaction."""
    html = "  \n  <div>  \n    <p>Hello</p>  \n  </div>  \n  "
    compacted = compact_html(html)
    assert compacted == "<div>\n<p>Hello</p>\n</div>"

# === TDD测试：提升HTMLUtils覆盖率到100% ===

def test_escape_html_with_none_input():
    """
    TDD测试：escape_html应该处理None输入

    这个测试覆盖第20行的None处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = escape_html(None)
    assert result == ""

def test_escape_html_with_empty_string():
    """
    TDD测试：escape_html应该处理空字符串

    这个测试确保空字符串被正确处理
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = escape_html("")
    assert result == ""

def test_escape_html_with_no_special_characters():
    """
    TDD测试：escape_html应该保持普通文本不变

    这个测试确保没有特殊字符的文本不被修改
    """
    # 🔴 红阶段：编写测试描述期望的行为
    normal_text = "Hello World 123"
    result = escape_html(normal_text)
    assert result == normal_text

def test_generate_style_attribute_with_empty_list():
    """
    TDD测试：generate_style_attribute应该处理空样式列表

    这个测试覆盖第39行的空列表处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = generate_style_attribute([])
    assert result == ""

def test_generate_style_attribute_with_none_styles():
    """
    TDD测试：generate_style_attribute应该过滤None值

    这个测试确保None值被正确过滤
    """
    # 🔴 红阶段：编写测试描述期望的行为
    styles = ["color: red;", None, "font-size: 12px;", None]
    result = generate_style_attribute(styles)
    assert result == 'style="color: red; font-size: 12px;"'

def test_generate_class_attribute_with_empty_list():
    """
    TDD测试：generate_class_attribute应该处理空类列表

    这个测试覆盖第54行的空列表处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = generate_class_attribute([])
    assert result == ""

def test_generate_class_attribute_with_none_classes():
    """
    TDD测试：generate_class_attribute应该过滤None值

    这个测试确保None值被正确过滤
    """
    # 🔴 红阶段：编写测试描述期望的行为
    classes = ["class1", None, "class2", None]
    result = generate_class_attribute(classes)
    assert result == 'class="class1 class2"'

def test_create_html_element_with_minimal_parameters():
    """
    TDD测试：create_html_element应该处理最少参数

    这个测试确保方法在只有标签名时正确工作
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = create_html_element("div")
    assert result == "<div></div>"

def test_create_html_element_with_self_closing_tag():
    """
    TDD测试：create_html_element应该处理自闭合标签

    这个测试覆盖第119行的自闭合标签处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = create_html_element("br", self_closing=True)
    assert result == "<br />"

    # 测试带属性的自闭合标签
    result = create_html_element("img", attributes={"src": "image.jpg"}, self_closing=True)
    assert result == '<img src="image.jpg" />'

def test_create_table_cell_with_all_parameters():
    """
    TDD测试：create_table_cell应该处理所有参数

    这个测试确保方法能处理所有可能的参数组合
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = create_table_cell(
        content="Cell Content",
        is_header=True,
        colspan=3,
        rowspan=2,
        css_classes=["cell-class"],
        inline_styles={"color": "red"},
        attributes={"data-id": "123"}
    )

    assert result.startswith("<th")
    assert 'colspan="3"' in result
    assert 'rowspan="2"' in result
    assert 'class="cell-class"' in result
    assert 'style="color: red;"' in result
    assert 'data-id="123"' in result
    assert "Cell Content" in result

def test_create_table_cell_with_no_colspan_rowspan():
    """
    TDD测试：create_table_cell应该处理没有colspan和rowspan的情况

    这个测试确保方法在没有跨列跨行时正确工作
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = create_table_cell("Simple Cell")

    assert result.startswith("<td")
    assert 'colspan=' not in result
    assert 'rowspan=' not in result
    assert "Simple Cell" in result

def test_create_svg_element_with_attributes():
    """
    TDD测试：create_svg_element应该处理额外的属性

    这个测试确保SVG元素能包含额外的属性
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = create_svg_element(
        width=200,
        height=100,
        content="<rect />",
        attributes={"viewBox": "0 0 200 100", "class": "svg-chart"}
    )

    assert 'width="200px"' in result
    assert 'height="100px"' in result
    assert 'viewBox="0 0 200 100"' in result
    assert 'class="svg-chart"' in result
    assert "<rect />" in result

def test_compact_html_with_empty_string():
    """
    TDD测试：compact_html应该处理空字符串

    这个测试确保空字符串被正确处理
    """
    # 🔴 红阶段：编写测试描述期望的行为
    result = compact_html("")
    assert result == ""

def test_compact_html_with_no_whitespace():
    """
    TDD测试：compact_html应该保持已经紧凑的HTML不变

    这个测试确保已经紧凑的HTML不被过度处理
    """
    # 🔴 红阶段：编写测试描述期望的行为
    compact_html_input = "<div><p>Hello</p></div>"
    result = compact_html(compact_html_input)
    assert result == compact_html_input

