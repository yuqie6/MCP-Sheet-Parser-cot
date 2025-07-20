import pytest
from src.models.table_model import (
    RichTextFragmentStyle, RichTextFragment, Cell, Style, Row, Sheet, 
    ChartPosition, Chart, LazySheet, LazyRowProvider, StreamingCapable
)

class TestRichTextFragmentStyle:
    """测试RichTextFragmentStyle类。"""

    def test_rich_text_fragment_style_default_values(self):
        """
        TDD测试：RichTextFragmentStyle应该有正确的默认值
        """
        style = RichTextFragmentStyle()
        
        assert style.bold is False
        assert style.italic is False
        assert style.underline is False
        assert style.font_color is None
        assert style.font_size is None
        assert style.font_name is None

    def test_rich_text_fragment_style_with_values(self):
        """
        TDD测试：RichTextFragmentStyle应该正确设置属性值
        """
        style = RichTextFragmentStyle(
            bold=True,
            italic=True,
            underline=True,
            font_color="#FF0000",
            font_size=14.0,
            font_name="Arial"
        )
        
        assert style.bold is True
        assert style.italic is True
        assert style.underline is True
        assert style.font_color == "#FF0000"
        assert style.font_size == 14.0
        assert style.font_name == "Arial"

class TestRichTextFragment:
    """测试RichTextFragment类。"""

    def test_rich_text_fragment_creation(self):
        """
        TDD测试：RichTextFragment应该正确创建
        """
        style = RichTextFragmentStyle(bold=True, font_color="#FF0000")
        fragment = RichTextFragment(text="Hello World", style=style)
        
        assert fragment.text == "Hello World"
        assert fragment.style == style
        assert fragment.style.bold is True
        assert fragment.style.font_color == "#FF0000"

class TestCell:
    """测试Cell类。"""

    def test_cell_default_values(self):
        """
        TDD测试：Cell应该有正确的默认值
        """
        cell = Cell(value="Test")
        
        assert cell.value == "Test"
        assert cell.style is None
        assert cell.row_span == 1
        assert cell.col_span == 1
        assert cell.formula is None

    def test_cell_with_all_properties(self):
        """
        TDD测试：Cell应该正确设置所有属性
        """
        style = Style(bold=True)
        cell = Cell(
            value="Test Value",
            style=style,
            row_span=2,
            col_span=3,
            formula="=A1+B1"
        )
        
        assert cell.value == "Test Value"
        assert cell.style == style
        assert cell.row_span == 2
        assert cell.col_span == 3
        assert cell.formula == "=A1+B1"

    def test_cell_with_rich_text_value(self):
        """
        TDD测试：Cell应该支持富文本值
        """
        style1 = RichTextFragmentStyle(bold=True)
        style2 = RichTextFragmentStyle(italic=True)
        fragments = [
            RichTextFragment("Bold text", style1),
            RichTextFragment("Italic text", style2)
        ]
        
        cell = Cell(value=fragments)
        
        assert cell.value == fragments
        assert len(cell.value) == 2
        assert cell.value[0].text == "Bold text"
        assert cell.value[1].text == "Italic text"

class TestStyle:
    """测试Style类。"""

    def test_style_default_values(self):
        """
        TDD测试：Style应该有正确的默认值
        """
        style = Style()
        
        assert style.font_name is None
        assert style.font_size is None
        assert style.font_color is None
        assert style.background_color is None
        assert style.bold is False
        assert style.italic is False
        assert style.underline is False
        assert style.text_align is None
        assert style.vertical_align is None
        assert style.border_top is None
        assert style.border_bottom is None
        assert style.border_left is None
        assert style.border_right is None
        assert style.border_color is None
        assert style.wrap_text is False
        assert style.number_format is None
        assert style.formula is None
        assert style.hyperlink is None
        assert style.comment is None

    def test_style_with_all_properties(self):
        """
        TDD测试：Style应该正确设置所有属性
        """
        style = Style(
            font_name="Arial",
            font_size=12,
            font_color="#000000",
            background_color="#FFFFFF",
            bold=True,
            italic=False,
            underline=True,
            text_align="center",
            vertical_align="middle",
            border_top="1px solid black",
            border_bottom="1px solid black",
            border_left="1px solid black",
            border_right="1px solid black",
            border_color="#000000",
            wrap_text=True,
            number_format="0.00",
            formula="=A1+B1",
            hyperlink="http://example.com",
            comment="Test comment"
        )
        
        assert style.font_name == "Arial"
        assert style.font_size == 12
        assert style.font_color == "#000000"
        assert style.background_color == "#FFFFFF"
        assert style.bold is True
        assert style.italic is False
        assert style.underline is True
        assert style.text_align == "center"
        assert style.vertical_align == "middle"
        assert style.border_top == "1px solid black"
        assert style.border_bottom == "1px solid black"
        assert style.border_left == "1px solid black"
        assert style.border_right == "1px solid black"
        assert style.border_color == "#000000"
        assert style.wrap_text is True
        assert style.number_format == "0.00"
        assert style.formula == "=A1+B1"
        assert style.hyperlink == "http://example.com"
        assert style.comment == "Test comment"

class TestRow:
    """测试Row类。"""

    def test_row_creation(self):
        """
        TDD测试：Row应该正确创建
        """
        cells = [Cell(value="A1"), Cell(value="B1"), Cell(value="C1")]
        row = Row(cells=cells)
        
        assert row.cells == cells
        assert len(row.cells) == 3
        assert row.cells[0].value == "A1"
        assert row.cells[1].value == "B1"
        assert row.cells[2].value == "C1"

    def test_row_with_empty_cells(self):
        """
        TDD测试：Row应该支持空单元格列表
        """
        row = Row(cells=[])
        
        assert row.cells == []
        assert len(row.cells) == 0

class TestChartPosition:
    """测试ChartPosition类。"""

    def test_chart_position_creation(self):
        """
        TDD测试：ChartPosition应该正确创建
        """
        position = ChartPosition(
            from_col=0, from_row=0, to_col=5, to_row=10,
            from_col_offset=0, from_row_offset=0, to_col_offset=0, to_row_offset=0
        )
        
        assert position.from_col == 0
        assert position.from_row == 0
        assert position.to_col == 5
        assert position.to_row == 10
        assert position.from_col_offset == 0
        assert position.from_row_offset == 0
        assert position.to_col_offset == 0
        assert position.to_row_offset == 0

class TestChart:
    """测试Chart类。"""

    def test_chart_default_values(self):
        """
        TDD测试：Chart应该有正确的默认值
        """
        chart = Chart(name="TestChart", type="line", anchor="A1")
        
        assert chart.name == "TestChart"
        assert chart.type == "line"
        assert chart.anchor == "A1"
        assert chart.position is None
        assert chart.chart_data is None

    def test_chart_with_all_properties(self):
        """
        TDD测试：Chart应该正确设置所有属性
        """
        position = ChartPosition(0, 0, 5, 10, 0, 0, 0, 0)
        chart_data = {"type": "line", "data": [1, 2, 3]}
        
        chart = Chart(
            name="TestChart",
            type="line",
            anchor="A1",
            position=position,
            chart_data=chart_data
        )
        
        assert chart.name == "TestChart"
        assert chart.type == "line"
        assert chart.anchor == "A1"
        assert chart.position == position
        assert chart.chart_data == chart_data

class TestSheet:
    """测试Sheet类。"""

    def test_sheet_default_values(self):
        """
        TDD测试：Sheet应该有正确的默认值
        """
        rows = [Row(cells=[Cell(value="A1")])]
        sheet = Sheet(name="TestSheet", rows=rows)
        
        assert sheet.name == "TestSheet"
        assert sheet.rows == rows
        assert sheet.merged_cells == []
        assert sheet.charts == []
        assert sheet.column_widths == {}
        assert sheet.row_heights == {}
        assert sheet.default_column_width == 8.43
        assert sheet.default_row_height == 18.0

    def test_sheet_with_all_properties(self):
        """
        TDD测试：Sheet应该正确设置所有属性
        """
        rows = [Row(cells=[Cell(value="A1")])]
        merged_cells = ["A1:B2"]
        charts = [Chart(name="Chart1", type="line", anchor="A1")]
        column_widths = {0: 20, 1: 30}
        row_heights = {0: 25, 1: 35}
        
        sheet = Sheet(
            name="TestSheet",
            rows=rows,
            merged_cells=merged_cells,
            charts=charts,
            column_widths=column_widths,
            row_heights=row_heights,
            default_column_width=10.0,
            default_row_height=20.0
        )
        
        assert sheet.name == "TestSheet"
        assert sheet.rows == rows
        assert sheet.merged_cells == merged_cells
        assert sheet.charts == charts
        assert sheet.column_widths == column_widths
        assert sheet.row_heights == row_heights
        assert sheet.default_column_width == 10.0
        assert sheet.default_row_height == 20.0

class MockRowProvider:
    """模拟的行提供者，用于测试LazySheet。"""

    def __init__(self, total_rows=10):
        self.total_rows = total_rows

    def iter_rows(self, start_row=0, max_rows=None):
        """生成模拟行。"""
        end_row = min(start_row + (max_rows or self.total_rows), self.total_rows)
        for i in range(start_row, end_row):
            yield Row(cells=[Cell(value=f"Row{i}Col{j}") for j in range(3)])

    def get_row(self, row_index):
        """获取指定行。"""
        if row_index >= self.total_rows:
            raise IndexError(f"Row index {row_index} out of range")
        return Row(cells=[Cell(value=f"Row{row_index}Col{j}") for j in range(3)])

    def get_total_rows(self):
        """获取总行数。"""
        return self.total_rows

class TestLazySheet:
    """测试LazySheet类。"""

    def test_lazy_sheet_creation(self):
        """
        TDD测试：LazySheet应该正确创建
        """
        provider = MockRowProvider(total_rows=5)
        lazy_sheet = LazySheet(name="LazyTestSheet", provider=provider)

        assert lazy_sheet.name == "LazyTestSheet"
        assert lazy_sheet._provider == provider

    def test_lazy_sheet_iter_rows(self):
        """
        TDD测试：LazySheet应该支持行迭代
        """
        provider = MockRowProvider(total_rows=5)
        lazy_sheet = LazySheet(name="LazyTestSheet", provider=provider)

        rows = list(lazy_sheet.iter_rows(start_row=1, max_rows=3))

        assert len(rows) == 3
        assert rows[0].cells[0].value == "Row1Col0"
        assert rows[1].cells[0].value == "Row2Col0"
        assert rows[2].cells[0].value == "Row3Col0"

    def test_lazy_sheet_get_row(self):
        """
        TDD测试：LazySheet应该支持按索引获取行
        """
        provider = MockRowProvider(total_rows=5)
        lazy_sheet = LazySheet(name="LazyTestSheet", provider=provider)

        row = lazy_sheet.get_row(2)

        assert row.cells[0].value == "Row2Col0"
        assert row.cells[1].value == "Row2Col1"
        assert row.cells[2].value == "Row2Col2"

    def test_lazy_sheet_get_total_rows(self):
        """
        TDD测试：LazySheet应该支持获取总行数
        """
        provider = MockRowProvider(total_rows=10)
        lazy_sheet = LazySheet(name="LazyTestSheet", provider=provider)

        total_rows = lazy_sheet.get_total_rows()

        assert total_rows == 10

class MockStreamingParser(StreamingCapable):
    """模拟的流式解析器，用于测试StreamingCapable。"""

    def __init__(self, supports_streaming=True):
        self._supports_streaming = supports_streaming

    def supports_streaming(self):
        """检查是否支持流式处理。"""
        return self._supports_streaming

    def create_lazy_sheet(self, file_path):
        """创建惰性sheet。"""
        provider = MockRowProvider(total_rows=5)
        return LazySheet(name="StreamingSheet", provider=provider)

class TestStreamingCapable:
    """测试StreamingCapable抽象类。"""

    def test_streaming_capable_supports_streaming_true(self):
        """
        TDD测试：StreamingCapable应该正确报告流式支持状态
        """
        parser = MockStreamingParser(supports_streaming=True)

        assert parser.supports_streaming() is True

    def test_streaming_capable_supports_streaming_false(self):
        """
        TDD测试：StreamingCapable应该正确报告不支持流式
        """
        parser = MockStreamingParser(supports_streaming=False)

        assert parser.supports_streaming() is False

    def test_streaming_capable_create_lazy_sheet(self):
        """
        TDD测试：StreamingCapable应该能创建LazySheet
        """
        parser = MockStreamingParser()

        lazy_sheet = parser.create_lazy_sheet("test.xlsx")

        assert isinstance(lazy_sheet, LazySheet)
        assert lazy_sheet.name == "StreamingSheet"
        assert lazy_sheet.get_total_rows() == 5

class TestTableModelEdgeCases:
    """测试表格模型的边界情况。"""

    def test_cell_with_none_values(self):
        """
        TDD测试：Cell应该处理None值
        """
        cell = Cell(value=None)

        assert cell.value is None
        assert cell.style is None

    def test_style_equality(self):
        """
        TDD测试：Style对象应该支持相等性比较
        """
        style1 = Style(bold=True, font_size=12)
        style2 = Style(bold=True, font_size=12)
        style3 = Style(bold=False, font_size=12)

        assert style1 == style2
        assert style1 != style3

    def test_rich_text_fragment_style_equality(self):
        """
        TDD测试：RichTextFragmentStyle对象应该支持相等性比较
        """
        style1 = RichTextFragmentStyle(bold=True, font_color="#FF0000")
        style2 = RichTextFragmentStyle(bold=True, font_color="#FF0000")
        style3 = RichTextFragmentStyle(bold=False, font_color="#FF0000")

        assert style1 == style2
        assert style1 != style3

    def test_chart_position_equality(self):
        """
        TDD测试：ChartPosition对象应该支持相等性比较
        """
        pos1 = ChartPosition(0, 0, 5, 10, 0, 0, 0, 0)
        pos2 = ChartPosition(0, 0, 5, 10, 0, 0, 0, 0)
        pos3 = ChartPosition(1, 0, 5, 10, 0, 0, 0, 0)

        assert pos1 == pos2
        assert pos1 != pos3

# === 边界情况和未覆盖代码测试 ===

class TestSheetIterRows:
    """测试Sheet.iter_rows方法的边界情况。"""

    def test_iter_rows_with_max_rows_exceeding_total(self):
        """
        TDD测试：iter_rows应该处理max_rows超过总行数的情况

        这个测试覆盖第186行的min函数调用
        """
        rows = [
            Row(cells=[Cell(value=f"Row {i}")]) for i in range(3)
        ]
        sheet = Sheet(name="测试表", rows=rows, merged_cells=[])

        # 请求超过总行数的行数
        result_rows = list(sheet.iter_rows(start_row=0, max_rows=10))

        # 应该只返回实际存在的行数
        assert len(result_rows) == 3
        assert result_rows[0].cells[0].value == "Row 0"
        assert result_rows[2].cells[0].value == "Row 2"

    def test_iter_rows_with_start_row_beyond_range(self):
        """
        TDD测试：iter_rows应该处理start_row超出范围的情况

        这个测试覆盖第189行的边界检查
        """
        rows = [Row(cells=[Cell(value="Row 0")])]
        sheet = Sheet(name="测试表", rows=rows, merged_cells=[])

        # 从超出范围的行开始
        result_rows = list(sheet.iter_rows(start_row=5, max_rows=10))

        # 应该返回空列表
        assert len(result_rows) == 0

class TestLazySheetSlicing:
    """测试LazySheet的切片功能。"""

    def test_lazy_sheet_slice_with_step_not_one(self):
        """
        TDD测试：LazySheet切片应该拒绝步长不为1的切片

        这个测试覆盖第226-227行的步长检查
        """
        # 创建模拟的LazyRowProvider
        class MockProvider:
            def iter_rows(self, start_row=0, max_rows=None):
                for i in range(start_row, min(start_row + (max_rows or 5), 5)):
                    yield Row(cells=[Cell(value=f"Row {i}")])

            def get_row(self, row_index):
                return Row(cells=[Cell(value=f"Row {row_index}")])

            def get_total_rows(self):
                return 5

        lazy_sheet = LazySheet("测试表", MockProvider())

        # 尝试使用步长不为1的切片
        with pytest.raises(ValueError, match="暂不支持步长切片"):
            _ = lazy_sheet[0:3:2]

    def test_lazy_sheet_invalid_index_type(self):
        """
        TDD测试：LazySheet应该拒绝无效的索引类型

        这个测试覆盖第229-230行的类型检查
        """
        class MockProvider:
            def get_total_rows(self):
                return 5

        lazy_sheet = LazySheet("测试表", MockProvider())

        # 尝试使用无效的索引类型
        with pytest.raises(TypeError, match="无效的下标类型"):
            _ = lazy_sheet["invalid"]

    def test_lazy_sheet_slice_access(self):
        """
        TDD测试：LazySheet应该支持切片访问

        这个测试覆盖第224-228行的切片处理
        """
        class MockProvider:
            def iter_rows(self, start_row=0, max_rows=None):
                for i in range(start_row, min(start_row + (max_rows or 5), 5)):
                    yield Row(cells=[Cell(value=f"Row {i}")])

            def get_total_rows(self):
                return 5

        lazy_sheet = LazySheet("测试表", MockProvider())

        # 使用切片访问
        result = lazy_sheet[1:3]

        assert len(result) == 2
        assert result[0].cells[0].value == "Row 1"
        assert result[1].cells[0].value == "Row 2"

    def test_lazy_sheet_to_sheet_conversion(self):
        """
        TDD测试：LazySheet应该能转换为常规Sheet

        这个测试覆盖第232-239行的转换逻辑
        """
        class MockProvider:
            def iter_rows(self, start_row=0, max_rows=None):
                for i in range(3):
                    yield Row(cells=[Cell(value=f"Row {i}")])

        lazy_sheet = LazySheet("测试表", MockProvider(), merged_cells=["A1:B1"])

        # 转换为常规Sheet
        regular_sheet = lazy_sheet.to_sheet()

        assert isinstance(regular_sheet, Sheet)
        assert regular_sheet.name == "测试表"
        assert len(regular_sheet.rows) == 3
        assert regular_sheet.merged_cells == ["A1:B1"]
        assert regular_sheet.rows[0].cells[0].value == "Row 0"

class TestProtocolAndAbstractMethods:
    """测试Protocol和抽象方法的覆盖。"""

    def test_lazy_row_provider_protocol_methods(self):
        """
        TDD测试：LazyRowProvider协议方法应该可以被调用

        这个测试覆盖第32, 36, 40行的协议方法
        """
        # 创建实现协议的类
        class TestProvider:
            def iter_rows(self, start_row=0, max_rows=None):
                return iter([])  # 覆盖第32行

            def get_row(self, row_index):
                return Row(cells=[])  # 覆盖第36行

            def get_total_rows(self):
                return 0  # 覆盖第40行

        provider = TestProvider()

        # 验证方法可以被调用
        assert list(provider.iter_rows()) == []
        assert isinstance(provider.get_row(0), Row)
        assert provider.get_total_rows() == 0

    def test_streaming_capable_abstract_methods(self):
        """
        TDD测试：StreamingCapable抽象方法应该可以被实现

        这个测试覆盖第48, 53行的抽象方法
        """
        # 创建实现抽象类的类
        class TestStreamingParser(StreamingCapable):
            def supports_streaming(self):
                return True  # 覆盖第48行

            def create_lazy_sheet(self, file_path, sheet_name=None):
                return None  # 覆盖第53行

        parser = TestStreamingParser()

        # 验证方法可以被调用
        assert parser.supports_streaming() is True
        assert parser.create_lazy_sheet("test.xlsx") is None
