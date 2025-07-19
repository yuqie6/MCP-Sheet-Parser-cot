

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

