import pytest
from dataclasses import is_dataclass
from src.models.table_model import Style, Cell, Row, Sheet

 # 标记本模块所有测试为 'models'
pytestmark = pytest.mark.models

class TestStyleModel:
    def test_style_is_dataclass(self):
        # Style 应为 dataclass
        assert is_dataclass(Style)

    def test_style_defaults(self):
        # Style 默认值测试
        style = Style()
        assert not style.bold
        assert not style.italic
        assert style.font_color == "#000000"
        assert style.background_color == "#FFFFFF"

    def test_style_custom_values(self):
        # Style 自定义值测试
        style = Style(bold=True, font_color="#FF0000")
        assert style.bold
        assert style.font_color == "#FF0000"

    def test_style_advanced_features(self):
        # Style 进阶功能测试
        style = Style(
            hyperlink="https://example.com",
            comment="This is a test comment"
        )
        assert style.hyperlink == "https://example.com"
        assert style.comment == "This is a test comment"

        # 测试默认值
        default_style = Style()
        assert default_style.hyperlink is None
        assert default_style.comment is None

class TestCellModel:
    def test_cell_is_dataclass(self):
        # Cell 应为 dataclass
        assert is_dataclass(Cell)

    def test_cell_minimal_creation(self):
        # Cell 最简创建
        cell = Cell(value="Test")
        assert cell.value == "Test"
        assert cell.style is None
        assert cell.row_span == 1
        assert cell.col_span == 1

    def test_cell_with_style_and_span(self):
        # Cell 带样式和跨行跨列
        style = Style(bold=True)
        cell = Cell(value=123, style=style, row_span=2, col_span=3)
        assert cell.value == 123
        assert cell.style.bold
        assert cell.row_span == 2
        assert cell.col_span == 3

class TestRowModel:
    def test_row_is_dataclass(self):
        # Row 应为 dataclass
        assert is_dataclass(Row)

    def test_row_creation(self):
        # Row 创建测试
        cell1 = Cell(value="A1")
        cell2 = Cell(value="B1")
        row = Row(cells=[cell1, cell2])
        assert len(row.cells) == 2
        assert row.cells[0].value == "A1"

    def test_row_empty_creation(self):
        # Row 空行创建
        row = Row(cells=[])
        assert len(row.cells) == 0

class TestSheetModel:
    def test_sheet_is_dataclass(self):
        # Sheet 应为 dataclass
        assert is_dataclass(Sheet)

    def test_sheet_creation(self):
        # Sheet 创建测试
        row1 = Row(cells=[Cell("A1"), Cell("B1")])
        row2 = Row(cells=[Cell("A2"), Cell("B2")])
        sheet = Sheet(name="TestSheet", rows=[row1, row2], merged_cells=["A1:B1"])
        
        assert sheet.name == "TestSheet"
        assert len(sheet.rows) == 2
        assert len(sheet.merged_cells) == 1
        assert sheet.merged_cells[0] == "A1:B1"

    def test_sheet_defaults(self):
        # Sheet 默认值测试
        sheet = Sheet(name="DefaultSheet", rows=[])
        assert sheet.name == "DefaultSheet"
        assert len(sheet.rows) == 0
        assert isinstance(sheet.merged_cells, list)
        assert len(sheet.merged_cells) == 0
