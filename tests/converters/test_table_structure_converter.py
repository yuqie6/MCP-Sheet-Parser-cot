
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

# === TDD测试：提升TableStructureConverter覆盖率 ===

def test_generate_table_with_empty_sheet():
    """
    TDD测试：generate_table应该处理空工作表

    这个测试覆盖第22-34行的空工作表处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    empty_sheet = Sheet(name="EmptySheet", rows=[])

    html = converter.generate_table(empty_sheet, styles={}, header_rows=0)

    # 应该返回包含空表格的HTML
    assert '<table role="table"' in html
    assert '<tbody>' in html
    assert '</tbody>' in html

def test_generate_table_with_no_header_rows():
    """
    TDD测试：generate_table应该处理header_rows=0的情况

    这个测试覆盖第78行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "converted_content"
    style_converter = MagicMock()
    style_converter.get_style_key.return_value = "style_key"
    converter = TableStructureConverter(cell_converter, style_converter)

    row = Row(cells=[Cell(value="Data1"), Cell(value="Data2")])
    sheet = Sheet(name="TestSheet", rows=[row])

    html = converter.generate_table(sheet, styles={}, header_rows=0)

    # 应该没有thead，所有行都在tbody中
    assert '<thead>' not in html
    assert '<tbody>' in html
    assert '<td>converted_content</td>' in html

def test_generate_table_with_header_rows_exceeding_total():
    """
    TDD测试：generate_table应该处理header_rows超过总行数的情况

    这个测试覆盖第82行的边界条件
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "header_content"
    style_converter = MagicMock()
    style_converter.get_style_key.return_value = "style_key"
    converter = TableStructureConverter(cell_converter, style_converter)

    row = Row(cells=[Cell(value="Header1")])
    sheet = Sheet(name="TestSheet", rows=[row])

    html = converter.generate_table(sheet, styles={}, header_rows=5)  # 超过总行数

    # 所有行都应该被当作header
    assert '<thead>' in html
    assert '<th>header_content</th>' in html
    assert '<tbody>' in html  # 应该有空的tbody

def test_generate_row_html_with_header():
    """
    TDD测试：_generate_row_html应该正确生成header行

    这个测试覆盖第92行的is_header=True代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "header_cell"
    style_converter = MagicMock()
    style_converter.get_style_key.return_value = "header_style"
    converter = TableStructureConverter(cell_converter, style_converter)

    row = Row(cells=[Cell(value="Header")])
    styles = {"header_style": "font-weight: bold;"}

    html = converter._generate_row_html(row, styles, is_header=True)

    # 应该生成th元素而不是td
    assert '<th' in html
    assert 'header_cell' in html
    assert '<td' not in html

def test_generate_row_html_with_data():
    """
    TDD测试：_generate_row_html应该正确生成数据行

    这个测试覆盖第96-98行的is_header=False代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "data_cell"
    style_converter = MagicMock()
    style_converter.get_style_key.return_value = "data_style"
    converter = TableStructureConverter(cell_converter, style_converter)

    row = Row(cells=[Cell(value="Data")])
    styles = {"data_style": "color: black;"}

    html = converter._generate_row_html(row, styles, is_header=False)

    # 应该生成td元素而不是th
    assert '<td' in html
    assert 'data_cell' in html
    assert '<th' not in html

def test_generate_cell_html_with_merged_cells():
    """
    TDD测试：_generate_cell_html应该处理合并单元格

    这个测试覆盖第105-109行的合并单元格处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "merged_content"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    cell = Cell(value="Merged Cell")

    # 测试colspan
    html = converter._generate_cell_html(cell, "style_class", "style_attr", False, colspan=3)
    assert 'colspan="3"' in html

    # 测试rowspan
    html = converter._generate_cell_html(cell, "style_class", "style_attr", False, rowspan=2)
    assert 'rowspan="2"' in html

    # 测试同时有colspan和rowspan
    html = converter._generate_cell_html(cell, "style_class", "style_attr", False, colspan=2, rowspan=3)
    assert 'colspan="2"' in html
    assert 'rowspan="3"' in html

def test_generate_cell_html_with_hyperlink_only():
    """
    TDD测试：_generate_cell_html应该处理只有超链接的单元格

    这个测试覆盖第134行的超链接处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "link_text"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    cell = Cell(value="Link Text", style=Style(hyperlink="https://example.com"))

    html = converter._generate_cell_html(cell, "", "", False)

    # 应该包含超链接但没有title属性
    assert '<a href="https://example.com">' in html
    assert 'link_text' in html
    assert 'title=' not in html

def test_generate_cell_html_with_comment_only():
    """
    TDD测试：_generate_cell_html应该处理只有注释的单元格

    这个测试覆盖第140行的注释处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "comment_text"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    cell = Cell(value="Text with comment", style=Style(comment="This is a comment"))

    html = converter._generate_cell_html(cell, "", "", False)

    # 应该包含title属性但没有超链接
    assert 'title="This is a comment"' in html
    assert '<a href=' not in html

def test_generate_cell_html_with_formula_only():
    """
    TDD测试：_generate_cell_html应该处理只有公式的单元格

    这个测试覆盖第142行的公式处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "formula_result"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    cell = Cell(value="100", formula="=SUM(A1:A10)")

    html = converter._generate_cell_html(cell, "", "", False)

    # 应该包含公式信息在title中
    assert 'title="Formula: =SUM(A1:A10)"' in html
    assert 'formula_result' in html
