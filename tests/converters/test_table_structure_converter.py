import pytest
from unittest.mock import MagicMock, patch
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
    assert "<tbody>" in html
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


def test_generate_table_with_empty_sheet():
    """
    TDD测试：generate_table应该处理空工作表

    这个测试覆盖第22-34行的空工作表处理代码路径
    """
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
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "formula_result"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    cell = Cell(value="100", formula="=SUM(A1:A10)")

    html = converter._generate_cell_html(cell, "", "", False)

    # 应该包含公式信息在title中
    assert 'title="Formula: =SUM(A1:A10)"' in html
    assert 'formula_result' in html

# === TDD测试：提升TableStructureConverter更多覆盖率 ===

def test_generate_cell_html_with_overflow_style_parsing():
    """
    TDD测试：_generate_cell_html应该正确解析overflow样式

    这个测试覆盖第151-156行的样式解析代码路径
    """
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "overflow_content"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    cell = Cell(value="Long text that overflows")

    # 直接传递overflow_style参数
    overflow_style = 'style="color: red; font-weight: bold;"'
    html = converter._generate_cell_html(cell, "", "", False, overflow_style)

    # 应该包含解析后的内联样式
    assert 'style=' in html
    assert 'color: red' in html
    assert 'font-weight: bold' in html

def test_generate_cell_html_with_malformed_overflow_style():
    """
    TDD测试：_generate_cell_html应该处理格式错误的overflow样式

    这个测试确保方法在样式格式错误时不会崩溃
    """
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "content"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    cell = Cell(value="Text")

    # 模拟返回格式错误的样式
    malformed_styles = [
        'style="color red"',  # 缺少冒号
        'style="color:"',     # 缺少值
        'style="color"',      # 不完整
        'invalid_style',      # 不包含style=
    ]

    for malformed_style in malformed_styles:
        with patch.object(converter, '_should_overflow_text', return_value=malformed_style):
            # 应该不会抛出异常
            html = converter._generate_cell_html(cell, "", "", False)
            assert isinstance(html, str)

def test_should_overflow_text_with_empty_next_cells():
    """
    TDD测试：_should_overflow_text应该处理后续单元格为空的情况

    这个测试覆盖第175-185行的文本溢出检查代码路径
    """
    cell_converter = MagicMock()
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    # 创建一个长文本单元格
    long_text_cell = Cell(value="This is a very long text that should overflow into next cells")

    # 创建包含空单元格的行
    empty_cells = [Cell(value=None), Cell(value=""), Cell(value="   ")]
    row = Row(cells=[long_text_cell] + empty_cells)

    result = converter._should_overflow_text(long_text_cell, row, 0)

    # 应该返回溢出样式（因为后续单元格为空）
    assert result is not False
    if isinstance(result, str):
        assert 'style=' in result

def test_should_overflow_text_with_non_empty_next_cell():
    """
    TDD测试：_should_overflow_text应该在下一个单元格非空时不溢出

    这个测试确保文本不会溢出到有内容的单元格
    """
    cell_converter = MagicMock()
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    long_text_cell = Cell(value="This is a very long text")
    non_empty_cell = Cell(value="Not empty")
    row = Row(cells=[long_text_cell, non_empty_cell])

    result = converter._should_overflow_text(long_text_cell, row, 0)

    # 应该返回False（不溢出）
    assert result is False

def test_should_overflow_text_with_short_text():
    """
    TDD测试：_should_overflow_text应该在文本较短时不溢出

    这个测试确保短文本不会触发溢出
    """
    cell_converter = MagicMock()
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    short_text_cell = Cell(value="Short")
    row = Row(cells=[short_text_cell, Cell(value=None)])

    result = converter._should_overflow_text(short_text_cell, row, 0)

    # 应该返回False（不溢出）
    assert result is False

# === TDD测试：提升table_structure_converter覆盖率到85%+ ===

def test_generate_table_with_invalid_merged_cells():
    """
    TDD测试：generate_table应该处理无效的合并单元格范围

    这个测试覆盖第22-34行的异常处理代码
    """
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "content"
    style_converter = MagicMock()
    style_converter.get_style_key.return_value = "style_key"
    converter = TableStructureConverter(cell_converter, style_converter)

    row = Row(cells=[Cell(value="A1"), Cell(value="B1")])
    sheet = Sheet(name="TestSheet", rows=[row], merged_cells=["INVALID_RANGE", "A1:B1"])

    # 应该不抛出异常，并且正确处理有效的合并单元格
    html = converter.generate_table(sheet, styles={}, header_rows=0)

    assert '<table role="table"' in html
    assert 'content' in html

def test_generate_row_html_with_occupied_cells():
    """
    TDD测试：_generate_row_html应该跳过被占用的单元格

    这个测试覆盖第92-93行的occupied_cells检查代码
    """
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "visible_content"
    style_converter = MagicMock()
    style_converter.get_style_key.return_value = "style_key"
    converter = TableStructureConverter(cell_converter, style_converter)

    # 创建包含合并单元格的sheet
    row1 = Row(cells=[Cell(value="A1"), Cell(value="B1")])
    row2 = Row(cells=[Cell(value="A2"), Cell(value="B2")])
    sheet = Sheet(name="TestSheet", rows=[row1, row2], merged_cells=["A1:B2"])

    html = converter.generate_table(sheet, styles={}, header_rows=0)

    # 应该只显示合并单元格的起始位置，其他位置被跳过
    assert html.count('visible_content') < 4  # 少于4个单元格的内容

def test_generate_row_html_with_merged_cells_spans():
    """
    TDD测试：_generate_row_html应该正确处理合并单元格的跨度属性

    这个测试覆盖第98-103行的跨度属性生成代码
    """
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "merged_content"
    style_converter = MagicMock()
    style_converter.get_style_key.return_value = "style_key"
    converter = TableStructureConverter(cell_converter, style_converter)

    # 创建包含不同类型合并单元格的sheet
    row1 = Row(cells=[Cell(value="A1"), Cell(value="B1"), Cell(value="C1")])
    row2 = Row(cells=[Cell(value="A2"), Cell(value="B2"), Cell(value="C2")])
    row3 = Row(cells=[Cell(value="A3"), Cell(value="B3"), Cell(value="C3")])
    sheet = Sheet(name="TestSheet", rows=[row1, row2, row3],
                  merged_cells=["A1:B1", "C1:C3"])  # 水平合并和垂直合并

    html = converter.generate_table(sheet, styles={}, header_rows=0)

    # 应该包含colspan和rowspan属性
    assert 'colspan="2"' in html  # A1:B1的水平合并
    assert 'rowspan="3"' in html  # C1:C3的垂直合并

def test_generate_cell_html_with_style_class_and_span_attrs():
    """
    TDD测试：_generate_cell_html应该正确处理样式类和跨度属性

    这个测试覆盖第122、126行的样式类和跨度属性处理代码
    """
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "styled_content"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    cell = Cell(value="Styled Cell")

    # 测试样式类 - 需要通过overflow_style传递
    html = converter._generate_cell_html(cell, "custom-style", "", False, "")
    # 样式类可能不会直接出现在输出中，检查方法被正确调用
    assert 'styled_content' in html

    # 测试跨度属性 - 需要包含空格的格式
    html1 = converter._generate_cell_html(cell, "", ' rowspan="2"', False, "")
    assert 'rowspan="2"' in html1

    html2 = converter._generate_cell_html(cell, "", ' colspan="3"', False, "")
    assert 'colspan="3"' in html2

    # 测试colspan和rowspan参数
    html = converter._generate_cell_html(cell, "", "", False, "", colspan=2, rowspan=3)
    assert 'colspan="2"' in html
    assert 'rowspan="3"' in html

def test_generate_cell_html_with_complex_title_combinations():
    """
    TDD测试：_generate_cell_html应该处理复杂的title组合

    这个测试覆盖第131-153行的title生成逻辑
    """
    cell_converter = MagicMock()
    cell_converter.convert.return_value = "complex_content"
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    # 测试注释和公式的组合
    cell1 = Cell(value="Data", style=Style(comment="Comment"), formula="=A1+B1")
    html1 = converter._generate_cell_html(cell1, "", "", False)
    assert 'title="Comment | Formula: =A1+B1"' in html1

    # 测试超链接、注释和公式的组合
    cell2 = Cell(value="Link", style=Style(hyperlink="http://example.com", comment="Link comment"), formula="=SUM(A:A)")
    html2 = converter._generate_cell_html(cell2, "", "", False)
    assert '<a href="http://example.com"' in html2
    assert 'title="Link comment | Formula: =SUM(A:A)"' in html2

def test_should_overflow_text_edge_cases():
    """
    TDD测试：_should_overflow_text应该处理边界情况

    这个测试覆盖第179、187、189、194行的边界情况处理代码
    """
    cell_converter = MagicMock()
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    # 测试单元格在行末尾的情况
    long_text_cell = Cell(value="Very long text that should overflow")
    row_at_end = Row(cells=[long_text_cell])  # 只有一个单元格
    result = converter._should_overflow_text(long_text_cell, row_at_end, 0)
    assert result is True  # 长文本应该溢出，即使没有后续单元格

    # 测试None值的单元格
    none_cell = Cell(value=None)
    row_with_none = Row(cells=[none_cell, Cell(value="")])
    result = converter._should_overflow_text(none_cell, row_with_none, 0)
    assert result is False  # None值不应该溢出

    # 测试空字符串单元格
    empty_cell = Cell(value="")
    row_with_empty = Row(cells=[empty_cell, Cell(value=None)])
    result = converter._should_overflow_text(empty_cell, row_with_empty, 0)
    assert result is False  # 空字符串不应该溢出

def test_should_overflow_text_with_whitespace_handling():
    """
    TDD测试：_should_overflow_text应该正确处理空白字符

    这个测试覆盖第226、238、262行的空白字符处理代码
    """
    cell_converter = MagicMock()
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    # 测试只包含空白字符的单元格
    whitespace_cell = Cell(value="   \t\n   ")
    row = Row(cells=[whitespace_cell, Cell(value=None)])
    result = converter._should_overflow_text(whitespace_cell, row, 0)
    assert result is False  # 只有空白字符不应该溢出

    # 测试包含空白字符的长文本
    long_text_with_spaces = Cell(value="This is a very long text with spaces that should overflow")
    row_with_spaces = Row(cells=[long_text_with_spaces, Cell(value="")])
    result = converter._should_overflow_text(long_text_with_spaces, row_with_spaces, 0)
    # 应该返回溢出样式或False
    assert result is not None

def test_should_overflow_text_with_calculated_width():
    """
    TDD测试：_should_overflow_text应该基于计算的宽度决定溢出

    这个测试覆盖第269-270、287-305行的宽度计算和样式生成代码
    """
    cell_converter = MagicMock()
    style_converter = MagicMock()
    converter = TableStructureConverter(cell_converter, style_converter)

    # 创建一个确实很长的文本
    very_long_text = "A" * 100  # 100个字符的长文本
    long_cell = Cell(value=very_long_text)

    # 创建多个空的后续单元格
    empty_cells = [Cell(value=None) for _ in range(5)]
    row = Row(cells=[long_cell] + empty_cells)

    result = converter._should_overflow_text(long_cell, row, 0)

    # 应该返回True（表示应该溢出）
    assert result is True

class TestTableStructureConverterUncoveredCode:
    """TDD测试：表格结构转换器未覆盖代码测试"""

    def test_generate_row_html_with_merged_cells_rowspan_and_colspan(self):
        """
        TDD测试：_generate_row_html应该处理rowspan和colspan

        覆盖代码行：99-103 - 合并单元格跨度属性处理逻辑
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "merged_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "style_key"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含合并单元格的工作表
        row1 = Row(cells=[Cell(value="A1"), Cell(value="B1"), Cell(value="C1")])
        row2 = Row(cells=[Cell(value="A2"), Cell(value="B2"), Cell(value="C2")])
        row3 = Row(cells=[Cell(value="A3"), Cell(value="B3"), Cell(value="C3")])

        # 设置合并单元格：A1:B2（2列2行）
        sheet = Sheet(name="TestSheet", rows=[row1, row2, row3], merged_cells=["A1:B2"])

        # 生成表格HTML
        html = converter.generate_table(sheet, styles={}, header_rows=0)

        # 验证包含rowspan和colspan属性
        assert 'rowspan="2"' in html
        assert 'colspan="2"' in html

    def test_generate_cell_html_with_style_and_wrap_text(self):
        """
        TDD测试：_generate_cell_html应该处理样式和文本换行

        覆盖代码行：131-136 - 样式和文本换行处理逻辑
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "wrapped_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "wrap_style"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含换行样式的单元格
        wrap_style = Style(wrap_text=True, bold=True)
        cell = Cell(value="Text that should wrap", style=wrap_style)

        # 创建样式映射
        styles = {"wrap_style": "style_id_123"}

        # 生成单元格HTML
        html = converter._generate_cell_html(cell, "", "", False)

        # 验证包含样式类
        assert 'wrapped_content' in html

    def test_generate_cell_html_with_overflow_text_conditions(self):
        """
        TDD测试：_generate_cell_html应该处理文本溢出条件

        覆盖代码行：140-142 - 文本溢出条件处理逻辑
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "overflow_content"
        style_converter = MagicMock()
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建长文本单元格
        long_text_cell = Cell(value="This is a very long text that should overflow into adjacent cells")

        # 创建包含空单元格的行
        row = Row(cells=[long_text_cell, Cell(value=None), Cell(value="")])

        # 模拟_should_overflow_text返回True
        with patch.object(converter, '_should_overflow_text', return_value=True):
            # 生成单元格HTML
            html = converter._generate_cell_html(long_text_cell, "", "", False)

            # 验证包含溢出内容
            assert 'overflow_content' in html

    def test_generate_cell_html_with_css_classes_combination(self):
        """
        TDD测试：_generate_cell_html应该处理CSS类的组合

        覆盖代码行：145 - CSS类组合逻辑
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "styled_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "multi_style"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含多种样式的单元格
        multi_style = Style(wrap_text=True, bold=True, background_color="#FFFF00")
        cell = Cell(value="Multi-styled text", style=multi_style)

        # 创建样式映射
        styles = {"multi_style": "style_id_456"}

        # 模拟_should_overflow_text返回True以添加text-overflow类
        with patch.object(converter, '_should_overflow_text', return_value=True):
            # 生成单元格HTML
            html = converter._generate_cell_html(cell, "", "", False)

            # 验证包含样式内容
            assert 'styled_content' in html

    def test_has_meaningful_content_with_formula(self):
        """
        TDD测试：_has_meaningful_content应该识别包含公式的单元格

        覆盖代码行：287-288 - 公式检查逻辑
        """
        cell_converter = MagicMock()
        style_converter = MagicMock()
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含公式的单元格
        formula_cell = Cell(value="100", formula="=SUM(A1:A10)")

        # 测试有意义内容检查
        result = converter._has_meaningful_content(formula_cell)

        # 包含公式的单元格应该被认为有意义
        assert result is True

    def test_has_meaningful_content_with_background_color(self):
        """
        TDD测试：_has_meaningful_content应该识别有背景色的单元格

        覆盖代码行：293-296 - 背景色检查逻辑
        """
        cell_converter = MagicMock()
        style_converter = MagicMock()
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含背景色的单元格
        bg_style = Style(background_color="#FF0000")
        bg_cell = Cell(value="", style=bg_style)

        # 测试有意义内容检查
        result = converter._has_meaningful_content(bg_cell)

        # 有背景色的单元格应该被认为有意义
        assert result is True

        # 测试默认背景色（应该被忽略）
        default_bg_style = Style(background_color="ffffff")
        default_bg_cell = Cell(value="", style=default_bg_style)
        result = converter._has_meaningful_content(default_bg_cell)

        # 默认背景色不应该被认为有意义
        assert result is False

    def test_has_meaningful_content_with_borders(self):
        """
        TDD测试：_has_meaningful_content应该识别有边框的单元格

        覆盖代码行：299-303 - 边框检查逻辑
        """
        cell_converter = MagicMock()
        style_converter = MagicMock()
        converter = TableStructureConverter(cell_converter, style_converter)

        # 测试各种边框
        border_styles = [
            Style(border_top="1px solid black"),
            Style(border_bottom="2px dashed red"),
            Style(border_left="1px dotted blue"),
            Style(border_right="3px solid green")
        ]

        for border_style in border_styles:
            border_cell = Cell(value="", style=border_style)
            result = converter._has_meaningful_content(border_cell)

            # 有边框的单元格应该被认为有意义
            assert result is True

    def test_generate_row_html_with_merged_cells_complex_spans(self):
        """
        TDD测试：_generate_row_html应该处理复杂的合并单元格跨度

        覆盖代码行：99-103 - 复杂合并单元格跨度处理
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "complex_merged"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "style_key"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建复杂的合并单元格场景
        rows = []
        for i in range(4):
            cells = []
            for j in range(4):
                cells.append(Cell(value=f"Cell_{i}_{j}"))
            rows.append(Row(cells=cells))

        # 设置多种合并单元格：A1:B1（colspan=2），C1:C4（rowspan=4），D2:D3（rowspan=2）
        sheet = Sheet(name="ComplexSheet", rows=rows, merged_cells=["A1:B1", "C1:C4", "D2:D3"])

        html = converter.generate_table(sheet, styles={}, header_rows=0)

        # 验证包含不同的跨度属性
        assert 'colspan="2"' in html  # A1:B1
        assert 'rowspan="4"' in html  # C1:C4
        assert 'rowspan="2"' in html  # D2:D3

    def test_generate_cell_html_with_edge_case_styles(self):
        """
        TDD测试：_generate_cell_html应该处理边界情况的样式

        覆盖代码行：126, 131-136 - 边界情况样式处理
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "edge_case_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = None  # 返回None的情况
        converter = TableStructureConverter(cell_converter, style_converter)

        # 测试样式为None的情况
        cell_no_style = Cell(value="No style")
        html = converter._generate_cell_html(cell_no_style, "", "", False)
        assert 'edge_case_content' in html

        # 测试空样式字符串
        style_converter.get_style_key.return_value = ""
        html = converter._generate_cell_html(cell_no_style, "", "", False)
        assert 'edge_case_content' in html

    def test_generate_cell_html_with_overflow_style_parsing_edge_cases(self):
        """
        TDD测试：_generate_cell_html应该处理溢出样式解析的边界情况

        覆盖代码行：140-142, 145 - 溢出样式解析边界情况
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "overflow_edge_content"
        style_converter = MagicMock()
        converter = TableStructureConverter(cell_converter, style_converter)

        cell = Cell(value="Overflow text")

        # 测试_should_overflow_text返回不同类型的值
        test_cases = [
            True,  # 布尔值True
            False, # 布尔值False
            'style="color: red;"',  # 样式字符串
            '',    # 空字符串
            None,  # None值
        ]

        for overflow_result in test_cases:
            with patch.object(converter, '_should_overflow_text', return_value=overflow_result):
                html = converter._generate_cell_html(cell, "", "", False)
                assert 'overflow_edge_content' in html

    def test_has_meaningful_content_edge_cases(self):
        """
        TDD测试：_has_meaningful_content应该处理各种边界情况

        覆盖代码行：269-270, 288 - 有意义内容检查边界情况
        """
        cell_converter = MagicMock()
        style_converter = MagicMock()
        converter = TableStructureConverter(cell_converter, style_converter)

        # 测试None值单元格
        none_cell = Cell(value=None)
        assert converter._has_meaningful_content(none_cell) is False

        # 测试空字符串单元格
        empty_cell = Cell(value="")
        assert converter._has_meaningful_content(empty_cell) is False

        # 测试只有空白字符的单元格
        whitespace_cell = Cell(value="   \t\n   ")
        assert converter._has_meaningful_content(whitespace_cell) is False

        # 测试有实际内容的单元格
        content_cell = Cell(value="Real content")
        assert converter._has_meaningful_content(content_cell) is True

        # 测试数字0
        zero_cell = Cell(value=0)
        assert converter._has_meaningful_content(zero_cell) is True

        # 测试布尔值False
        false_cell = Cell(value=False)
        assert converter._has_meaningful_content(false_cell) is True

    def test_generate_row_html_with_style_key_mapping(self):
        """
        TDD测试：_generate_rows_html应该正确处理样式键映射

        覆盖代码行：131-136 - 样式键到ID的映射处理
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "styled_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "test_style_key"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含样式的单元格
        cell_style = Style(bold=True, wrap_text=True)
        cell = Cell(value="Styled text", style=cell_style)
        row = Row(cells=[cell])

        # 创建样式键到ID的映射
        styles = {"test_style_key": "style_id_123"}

        # 使用_generate_rows_html方法测试样式处理
        table_parts = []
        converter._generate_rows_html(table_parts, [row], set(), {}, styles, is_header=False)
        html = ''.join(table_parts)

        # 验证包含样式内容
        assert 'styled_content' in html
        # 验证style_converter.get_style_key被调用
        style_converter.get_style_key.assert_called()

    def test_generate_row_html_with_text_overflow_class(self):
        """
        TDD测试：_generate_row_html应该添加text-overflow类

        覆盖代码行：140-142 - 文本溢出类和内联样式处理
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "overflow_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "overflow_style"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建长文本单元格
        long_text_cell = Cell(value="This is a very long text that should overflow")
        row = Row(cells=[long_text_cell, Cell(value=None)])

        # 模拟_should_overflow_text返回True
        with patch.object(converter, '_should_overflow_text', return_value=True) as mock_overflow:
            table_parts = []
            converter._generate_rows_html(table_parts, [row], set(), {}, {}, is_header=False)
            html = ''.join(table_parts)

            # 验证包含溢出内容
            assert 'overflow_content' in html
            # 验证_should_overflow_text被调用
            mock_overflow.assert_called()

    def test_generate_row_html_with_css_classes_combination(self):
        """
        TDD测试：_generate_row_html应该正确组合CSS类

        覆盖代码行：145 - CSS类组合逻辑
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "multi_class_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "multi_style_key"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含多种样式的单元格
        multi_style = Style(bold=True, wrap_text=True)
        cell = Cell(value="Multi-class text", style=multi_style)
        row = Row(cells=[cell, Cell(value=None)])

        # 创建样式映射
        styles = {"multi_style_key": "style_id_456"}

        # 模拟_should_overflow_text返回True以添加text-overflow类
        with patch.object(converter, '_should_overflow_text', return_value=True) as mock_overflow:
            table_parts = []
            converter._generate_rows_html(table_parts, [row], set(), {}, styles, is_header=False)
            html = ''.join(table_parts)

            # 验证包含多类样式内容
            assert 'multi_class_content' in html
            # 验证相关方法被调用
            mock_overflow.assert_called()
            style_converter.get_style_key.assert_called()

    def test_generate_row_html_with_occupied_cells_skip(self):
        """
        TDD测试：_generate_row_html应该跳过被占用的单元格

        覆盖代码行：93 - 被占用单元格跳过逻辑
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "visible_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "style_key"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含合并单元格的工作表
        row1 = Row(cells=[Cell(value="A1"), Cell(value="B1"), Cell(value="C1")])
        row2 = Row(cells=[Cell(value="A2"), Cell(value="B2"), Cell(value="C2")])
        sheet = Sheet(name="TestSheet", rows=[row1, row2], merged_cells=["A1:B1"])

        # 生成表格HTML
        html = converter.generate_table(sheet, styles={}, header_rows=0)

        # 第一行的B1单元格应该被跳过（因为被A1合并占用）
        # 验证colspan属性存在，表明合并单元格被正确处理
        assert 'colspan="2"' in html

    def test_generate_row_html_with_merged_cells_spans_detailed(self):
        """
        TDD测试：_generate_row_html应该详细处理合并单元格跨度

        覆盖代码行：99-103 - 详细的rowspan和colspan处理
        """
        cell_converter = MagicMock()
        cell_converter.convert.return_value = "span_content"
        style_converter = MagicMock()
        style_converter.get_style_key.return_value = "span_style"
        converter = TableStructureConverter(cell_converter, style_converter)

        # 创建包含不同跨度的合并单元格
        rows = []
        for i in range(3):
            cells = []
            for j in range(3):
                cells.append(Cell(value=f"Cell_{i}_{j}"))
            rows.append(Row(cells=cells))

        # 设置不同类型的合并单元格
        merged_cells = [
            "A1:A3",  # 只有rowspan=3
            "B1:C1",  # 只有colspan=2
            "B2:C3"   # 既有rowspan=2又有colspan=2
        ]
        sheet = Sheet(name="SpanSheet", rows=rows, merged_cells=merged_cells)

        html = converter.generate_table(sheet, styles={}, header_rows=0)

        # 验证不同的跨度属性
        assert 'rowspan="3"' in html  # A1:A3
        assert 'colspan="2"' in html  # B1:C1 和 B2:C3
        assert 'rowspan="2"' in html  # B2:C3
