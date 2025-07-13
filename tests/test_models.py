"""
数据模型测试模块

全面测试所有数据模型类：Style、Cell、Row、Sheet
目标覆盖率：95%+
"""

import pytest
from src.models.table_model import Style, Cell, Row, Sheet


class TestStyle:
    """Style类的全面测试。"""
    
    def test_style_creation_default(self):
        """测试Style的默认创建。"""
        style = Style()

        # 验证所有默认值（根据实际的dataclass定义）
        assert style.bold is False
        assert style.italic is False
        assert style.underline is False
        assert style.font_size is None
        assert style.font_name is None
        assert style.font_color is None
        assert style.background_color is None
        assert style.text_align is None
        assert style.vertical_align is None
        assert style.wrap_text is False
        assert style.border_top is None
        assert style.border_bottom is None
        assert style.border_left is None
        assert style.border_right is None
        assert style.border_color is None
        assert style.number_format is None
        assert style.hyperlink is None
        assert style.comment is None
    
    def test_style_creation_with_values(self):
        """测试Style的带参数创建。"""
        style = Style(
            bold=True,
            italic=True,
            underline=True,
            font_size=14.0,
            font_name="Arial",
            font_color="#FF0000",
            background_color="#FFFF00",
            text_align="center",
            vertical_align="middle",
            wrap_text=True,
            border_top="1px solid",
            border_bottom="2px dashed",
            border_left="1px dotted",
            border_right="3px double",
            border_color="#000000",
            number_format="0.00",
            hyperlink="https://example.com",
            comment="测试注释"
        )
        
        # 验证所有设置的值
        assert style.bold is True
        assert style.italic is True
        assert style.underline is True
        assert style.font_size == 14.0
        assert style.font_name == "Arial"
        assert style.font_color == "#FF0000"
        assert style.background_color == "#FFFF00"
        assert style.text_align == "center"
        assert style.vertical_align == "middle"
        assert style.wrap_text is True
        assert style.border_top == "1px solid"
        assert style.border_bottom == "2px dashed"
        assert style.border_left == "1px dotted"
        assert style.border_right == "3px double"
        assert style.border_color == "#000000"
        assert style.number_format == "0.00"
        assert style.hyperlink == "https://example.com"
        assert style.comment == "测试注释"
    
    def test_style_equality(self):
        """测试Style的相等性比较。"""
        style1 = Style(bold=True, font_color="#FF0000")
        style2 = Style(bold=True, font_color="#FF0000")
        style3 = Style(bold=False, font_color="#FF0000")
        
        assert style1 == style2
        assert style1 != style3
    
    def test_style_repr(self):
        """测试Style的字符串表示。"""
        style = Style(bold=True, font_color="#FF0000")
        repr_str = repr(style)
        
        assert "Style" in repr_str
        assert "bold=True" in repr_str
        assert "#FF0000" in repr_str


class TestCell:
    """Cell类的全面测试。"""
    
    def test_cell_creation_default(self):
        """测试Cell的默认创建。"""
        # Cell需要value参数，使用None作为默认值
        cell = Cell(value=None)

        assert cell.value is None
        assert cell.style is None
        assert cell.row_span == 1  # 默认值
        assert cell.col_span == 1  # 默认值
    
    def test_cell_creation_with_value(self):
        """测试Cell的带值创建。"""
        cell = Cell(value="测试值")
        
        assert cell.value == "测试值"
        assert cell.style is None
    
    def test_cell_creation_with_style(self):
        """测试Cell的带样式创建。"""
        style = Style(bold=True)
        cell = Cell(value="测试值", style=style)
        
        assert cell.value == "测试值"
        assert cell.style == style
        assert cell.style.bold is True
    
    def test_cell_with_different_value_types(self):
        """测试Cell支持不同类型的值。"""
        # 字符串
        cell_str = Cell(value="文本")
        assert cell_str.value == "文本"
        
        # 数字
        cell_int = Cell(value=42)
        assert cell_int.value == 42
        
        cell_float = Cell(value=3.14)
        assert cell_float.value == 3.14
        
        # 布尔值
        cell_bool = Cell(value=True)
        assert cell_bool.value is True
        
        # None
        cell_none = Cell(value=None)
        assert cell_none.value is None
    
    def test_cell_equality(self):
        """测试Cell的相等性比较。"""
        style = Style(bold=True)
        
        cell1 = Cell(value="测试", style=style)
        cell2 = Cell(value="测试", style=style)
        cell3 = Cell(value="不同", style=style)
        cell4 = Cell(value="测试", style=None)
        
        assert cell1 == cell2
        assert cell1 != cell3
        assert cell1 != cell4
    
    def test_cell_repr(self):
        """测试Cell的字符串表示。"""
        cell = Cell(value="测试值")
        repr_str = repr(cell)
        
        assert "Cell" in repr_str
        assert "测试值" in repr_str


class TestRow:
    """Row类的全面测试。"""
    
    def test_row_creation_default(self):
        """测试Row的默认创建。"""
        # Row需要cells参数
        row = Row(cells=[])

        assert row.cells == []
    
    def test_row_creation_with_cells(self):
        """测试Row的带单元格创建。"""
        cells = [Cell(value="A"), Cell(value="B"), Cell(value="C")]
        row = Row(cells=cells)
        
        assert len(row.cells) == 3
        assert row.cells[0].value == "A"
        assert row.cells[1].value == "B"
        assert row.cells[2].value == "C"
    
    def test_row_add_cell(self):
        """测试向Row添加单元格。"""
        row = Row(cells=[])
        cell = Cell(value="新单元格")

        row.cells.append(cell)

        assert len(row.cells) == 1
        assert row.cells[0].value == "新单元格"
    
    def test_row_equality(self):
        """测试Row的相等性比较。"""
        cells1 = [Cell(value="A"), Cell(value="B")]
        cells2 = [Cell(value="A"), Cell(value="B")]
        cells3 = [Cell(value="A"), Cell(value="C")]
        
        row1 = Row(cells=cells1)
        row2 = Row(cells=cells2)
        row3 = Row(cells=cells3)
        
        assert row1 == row2
        assert row1 != row3
    
    def test_row_repr(self):
        """测试Row的字符串表示。"""
        cells = [Cell(value="A"), Cell(value="B")]
        row = Row(cells=cells)
        repr_str = repr(row)

        assert "Row" in repr_str
        # 检查包含单元格信息，而不是具体的数量格式
        assert "Cell" in repr_str


class TestSheet:
    """Sheet类的全面测试。"""
    
    def test_sheet_creation_default(self):
        """测试Sheet的默认创建。"""
        sheet = Sheet(name="测试表", rows=[])

        assert sheet.name == "测试表"
        assert sheet.rows == []
        assert sheet.merged_cells == []
    
    def test_sheet_creation_with_data(self):
        """测试Sheet的带数据创建。"""
        rows = [
            Row(cells=[Cell(value="A1"), Cell(value="B1")]),
            Row(cells=[Cell(value="A2"), Cell(value="B2")])
        ]
        merged_cells = ["A1:B1"]
        
        sheet = Sheet(name="数据表", rows=rows, merged_cells=merged_cells)
        
        assert sheet.name == "数据表"
        assert len(sheet.rows) == 2
        assert len(sheet.merged_cells) == 1
        assert sheet.merged_cells[0] == "A1:B1"
    
    def test_sheet_add_row(self):
        """测试向Sheet添加行。"""
        sheet = Sheet(name="测试表", rows=[])
        row = Row(cells=[Cell(value="新行")])

        sheet.rows.append(row)

        assert len(sheet.rows) == 1
        assert sheet.rows[0].cells[0].value == "新行"
    
    def test_sheet_get_cell_value(self):
        """测试获取Sheet中特定单元格的值。"""
        rows = [
            Row(cells=[Cell(value="A1"), Cell(value="B1")]),
            Row(cells=[Cell(value="A2"), Cell(value="B2")])
        ]
        sheet = Sheet(name="测试表", rows=rows)
        
        # 测试获取不同位置的单元格值
        assert sheet.rows[0].cells[0].value == "A1"
        assert sheet.rows[0].cells[1].value == "B1"
        assert sheet.rows[1].cells[0].value == "A2"
        assert sheet.rows[1].cells[1].value == "B2"
    
    def test_sheet_equality(self):
        """测试Sheet的相等性比较。"""
        rows1 = [Row(cells=[Cell(value="A1")])]
        rows2 = [Row(cells=[Cell(value="A1")])]
        rows3 = [Row(cells=[Cell(value="B1")])]
        
        sheet1 = Sheet(name="表1", rows=rows1)
        sheet2 = Sheet(name="表1", rows=rows2)
        sheet3 = Sheet(name="表1", rows=rows3)
        sheet4 = Sheet(name="表2", rows=rows1)
        
        assert sheet1 == sheet2
        assert sheet1 != sheet3
        assert sheet1 != sheet4
    
    def test_sheet_repr(self):
        """测试Sheet的字符串表示。"""
        rows = [Row(cells=[Cell(value="A1"), Cell(value="B1")])]
        sheet = Sheet(name="测试表", rows=rows)
        repr_str = repr(sheet)
        
        assert "Sheet" in repr_str
        assert "测试表" in repr_str
        assert "1" in repr_str  # 行数
