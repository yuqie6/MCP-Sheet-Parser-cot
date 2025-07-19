
import pytest
from unittest.mock import MagicMock
from src.converters.table_structure_converter import TableStructureConverter
from src.models.table_model import Sheet, Row, Cell, Style

@pytest.fixture
def table_structure_converter():
    """Fixture for TableStructureConverter."""
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "cell_content"
    style_converter = MagicMock()
    style_converter.get_style_key.return_value = "style_key"
    return TableStructureConverter(cell_converter, style_converter)

@pytest.fixture
def sample_sheet():
    """Fixture for a sample sheet."""
    row1 = Row(cells=[Cell(value="A1"), Cell(value="B1")])
    row2 = Row(cells=[Cell(value="A2"), Cell(value="B2")])
    return Sheet(name="TestSheet", rows=[row1, row2])

def test_generate_table(table_structure_converter, sample_sheet):
    """Test generating a simple table."""
    html = table_structure_converter.generate_table(sample_sheet, styles={}, header_rows=0)
    assert '<table role="table"' in html
    assert "<tbody>" not in html
    assert "<tr>" in html
    assert "<td>cell_content</td>" in html

def test_generate_table_with_header(table_structure_converter, sample_sheet):
    """Test generating a table with a header."""
    html = table_structure_converter.generate_table(sample_sheet, styles={}, header_rows=1)
    assert "<thead>" in html
    assert "<th>cell_content</th>" in html

def test_generate_cell_html(table_structure_converter):
    """Test generating HTML for a single cell."""
    cell = Cell(value="Data", style=Style(hyperlink="http://example.com", comment="My Comment"), formula="=SUM(A1:A2)")
    html = table_structure_converter._generate_cell_html(cell, "", "", False)
    assert '<a href="http://example.com">' in html
    assert 'title="My Comment | Formula: =SUM(A1:A2)"' in html

def test_should_overflow_text(table_structure_converter):
    """Test text overflow logic."""
    long_text_cell = Cell(value="This is a very long text that should overflow")
    short_text_cell = Cell(value="short")
    row = Row(cells=[long_text_cell, Cell(value=None)])
    assert table_structure_converter._should_overflow_text(long_text_cell, row, 0) is True
    assert table_structure_converter._should_overflow_text(short_text_cell, row, 0) is False
