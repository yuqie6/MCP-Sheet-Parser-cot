
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
