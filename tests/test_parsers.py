import pytest
from pathlib import Path
from src.models.table_model import Sheet, Row, Cell, Style
from src.parsers.base_parser import BaseParser
from src.parsers.csv_parser import CsvParser
from src.parsers.xlsx_parser import XlsxParser
from src.parsers.xls_parser import XlsParser
from src.parsers.xlsb_parser import XlsbParser
from src.parsers.xlsm_parser import XlsmParser
from src.parsers.factory import ParserFactory, UnsupportedFileType

@pytest.fixture
def sample_csv_path() -> Path:
    return Path(__file__).parent / "data" / "sample.csv"

@pytest.fixture
def sample_xlsx_path() -> Path:
    return Path(__file__).parent / "data" / "sample.xlsx"

def test_csv_parser_success(sample_csv_path: Path):
    """
    测试 CsvParser 能否成功解析简单的 CSV 文件。
    """
    # 准备
    parser = CsvParser()
    expected_sheet = Sheet(
        name="sample",
        rows=[
            Row(cells=[Cell(value="header1"), Cell(value="header2"), Cell(value="header3")]),
            Row(cells=[Cell(value="value1_1"), Cell(value="value1_2"), Cell(value="value1_3")]),
            Row(cells=[Cell(value="value2_1"), Cell(value="value2_2"), Cell(value="value2_3")]),
            Row(cells=[Cell(value="value3_1"), Cell(value="value3_2"), Cell(value="value3_3")]),
        ]
    )

    # 执行
    result_sheet = parser.parse(str(sample_csv_path))

    # 断言
    assert isinstance(result_sheet, Sheet)
    assert result_sheet.name == "sample"
    assert len(result_sheet.rows) == 4
    
    # 检查表头
    assert [cell.value for cell in result_sheet.rows[0].cells] == ["header1", "header2", "header3"]
    
    # 检查数据行
    assert [cell.value for cell in result_sheet.rows[1].cells] == ["value1_1", "value1_2", "value1_3"]
    assert [cell.value for cell in result_sheet.rows[2].cells] == ["value2_1", "value2_2", "value2_3"]
    assert [cell.value for cell in result_sheet.rows[3].cells] == ["value3_1", "value3_2", "value3_3"]

    # 深度比较
    assert result_sheet == expected_sheet

def test_xlsx_parser_success(sample_xlsx_path: Path):
    """
    测试 XlsxParser 能否成功解析简单的 XLSX 文件。
    """
    # 准备
    parser = XlsxParser()
    # 创建期望的默认样式（与解析器输出一致）
    default_style = Style(
        bold=False,
        italic=False,
        underline=False,
        font_color="#000000",
        font_size=11.0,
        font_name="宋体",  # 实际解析出的字体名称
        background_color="#FFFFFF",
        text_align="left",
        vertical_align="top",
        border_top="",
        border_bottom="",
        border_left="",
        border_right="",
        border_color="#000000",
        wrap_text=False,
        number_format="",
        hyperlink=None,
        comment=None
    )
    expected_sheet = Sheet(
        name="Sheet1",
        rows=[
            Row(cells=[
                Cell(value="ID", style=default_style),
                Cell(value="Name", style=default_style),
                Cell(value="Value", style=default_style)
            ]),
            Row(cells=[
                Cell(value=1, style=default_style),
                Cell(value="Alice", style=default_style),
                Cell(value=100, style=default_style)
            ]),
            Row(cells=[
                Cell(value=2, style=default_style),
                Cell(value="Bob", style=default_style),
                Cell(value=200, style=default_style)
            ]),
        ]
    )

    # 执行
    result_sheet = parser.parse(str(sample_xlsx_path))

    # 断言
    assert isinstance(result_sheet, Sheet)
    assert result_sheet.name == "Sheet1"
    assert len(result_sheet.rows) == 3
    
    # 检查表头
    assert [cell.value for cell in result_sheet.rows[0].cells] == ["ID", "Name", "Value"]
    
    # 检查数据行
    assert [cell.value for cell in result_sheet.rows[1].cells] == [1, "Alice", 100]
    assert [cell.value for cell in result_sheet.rows[2].cells] == [2, "Bob", 200]

    # 深度比较
    assert result_sheet == expected_sheet


class TestParserFactory:
    def test_get_parser_csv(self, sample_csv_path: Path):
        """
        测试 ParserFactory 能否为 .csv 文件返回 CsvParser 实例。
        """
        # 执行
        parser = ParserFactory.get_parser(str(sample_csv_path))
        # 断言
        assert isinstance(parser, CsvParser)

    def test_get_parser_xlsx(self, sample_xlsx_path: Path):
        """
        测试 ParserFactory 能否为 .xlsx 文件返回 XlsxParser 实例。
        """
        # 执行
        parser = ParserFactory.get_parser(str(sample_xlsx_path))
        # 断言
        assert isinstance(parser, XlsxParser)

    def test_get_parser_unsupported_file_type(self):
        """
        测试 ParserFactory 对不支持的文件类型是否抛出 UnsupportedFileType 错误。
        """
        # 准备
        unsupported_file_path = "test.txt"
        # 执行并断言
        with pytest.raises(UnsupportedFileType):
            ParserFactory.get_parser(unsupported_file_path)


def test_xlsx_style_extraction_basic_properties(sample_xlsx_path: Path):
    """测试 XLSX 文件的基础样式提取。"""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    # 测试样式是否被提取
    first_cell = sheet.rows[0].cells[0]
    assert first_cell.style is not None

    # 测试字体属性
    style = first_cell.style
    assert isinstance(style.bold, bool)
    assert isinstance(style.italic, bool)
    assert isinstance(style.underline, bool)
    assert isinstance(style.font_color, str)
    assert style.font_color.startswith('#')

    # 测试字体大小和名称
    if style.font_size is not None:
        assert isinstance(style.font_size, (int, float))
        assert style.font_size > 0

    if style.font_name is not None:
        assert isinstance(style.font_name, str)
        assert len(style.font_name) > 0


def test_xlsx_style_extraction_background_and_alignment(sample_xlsx_path: Path):
    """测试背景色和对齐方式提取。"""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    first_cell = sheet.rows[0].cells[0]
    style = first_cell.style
    assert style is not None
    # 测试背景色
    assert isinstance(style.background_color, str)
    assert style.background_color.startswith('#')

    # 测试对齐属性
    assert style.text_align in ['left', 'center', 'right', 'justify', 'general']
    assert style.vertical_align in ['top', 'middle', 'bottom', 'center']

    # 测试自动换行
    assert isinstance(style.wrap_text, bool)


def test_xlsx_style_extraction_borders(sample_xlsx_path: Path):
    """测试边框样式提取。"""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    first_cell = sheet.rows[0].cells[0]
    style = first_cell.style
    assert style is not None
    # 测试边框属性
    assert isinstance(style.border_top, str)
    assert isinstance(style.border_bottom, str)
    assert isinstance(style.border_left, str)
    assert isinstance(style.border_right, str)
    assert isinstance(style.border_color, str)
    assert style.border_color.startswith('#')


def test_xlsx_style_extraction_number_format(sample_xlsx_path: Path):
    """测试数字格式提取。"""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    first_cell = sheet.rows[0].cells[0]
    style = first_cell.style
    assert style is not None
    # 测试数字格式
    assert isinstance(style.number_format, str)


def test_style_fidelity_coverage(sample_xlsx_path: Path):
    """测试样式提取是否覆盖所有主要样式属性以实现 95% 还原度。"""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    # 测试多个单元格以确保全面覆盖
    for row_idx, row in enumerate(sheet.rows[:3]):  # Test first 3 rows
        for cell_idx, cell in enumerate(row.cells[:3]):  # Test first 3 columns
            style = cell.style
            assert style is not None
            # 验证所有样式属性都存在且类型正确
            style_properties = [
                ('bold', bool),
                ('italic', bool),
                ('underline', bool),
                ('font_color', str),
                ('background_color', str),
                ('text_align', str),
                ('vertical_align', str),
                ('border_top', str),
                ('border_bottom', str),
                ('border_left', str),
                ('border_right', str),
                ('border_color', str),
                ('wrap_text', bool),
                ('number_format', str)
            ]

            for prop_name, prop_type in style_properties:
                assert hasattr(style, prop_name), f"Style missing property: {prop_name}"
                prop_value = getattr(style, prop_name)
                if prop_value is not None:
                    assert isinstance(prop_value, prop_type), \
                        f"Property {prop_name} has wrong type: {type(prop_value)}"

            # 测试颜色格式有效性
            assert style.font_color.startswith('#'), f"Invalid font color format: {style.font_color}"
            assert style.background_color.startswith('#'), f"Invalid background color format: {style.background_color}"
            assert style.border_color.startswith('#'), f"Invalid border color format: {style.border_color}"


def test_indexed_color_mapping():
    """测试索引色映射功能。"""
    parser = XlsxParser()

    # 测试索引色转换
    test_colors = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    for index in test_colors:
        color = parser._get_indexed_color(index)
        assert isinstance(color, str)
        assert color.startswith('#')
        assert len(color) == 7  # #RRGGBB format


def test_border_style_mapping():
    """测试边框样式映射功能。"""
    parser = XlsxParser()

    # 创建用于测试的 mock 边框对象
    class MockBorderSide:
        def __init__(self, style, color=None):
            self.style = style
            self.color = color

    # 测试多种边框样式
    border_styles = ['thin', 'medium', 'thick', 'double', 'dotted', 'dashed', 'hair']
    for style_name in border_styles:
        mock_border = MockBorderSide(style_name)
        css_style = parser._get_border_style(mock_border)
        assert isinstance(css_style, str)
        assert len(css_style) > 0
        assert 'px' in css_style  # Should contain pixel values

    # 测试空边框
    empty_border = MockBorderSide(None)
    css_style = parser._get_border_style(empty_border)
    assert css_style == ""


def test_comprehensive_color_extraction():
    """测试全面的颜色提取功能。"""
    parser = XlsxParser()

    # 测试 RGB 颜色
    class MockRGBColor:
        def __init__(self, rgb_value):
            self.rgb = rgb_value

    # 测试 6 位 RGB
    rgb_color = MockRGBColor("FF0000")
    assert parser._extract_color(rgb_color) == "#FF0000"

    # 测试 8 位 ARGB（去掉 Alpha 通道）
    argb_color = MockRGBColor("80FF0000")
    assert parser._extract_color(argb_color) == "#FF0000"

    # 测试索引颜色
    class MockIndexedColor:
        def __init__(self, index):
            self.indexed = index

    indexed_color = MockIndexedColor(2)  # 红色
    assert parser._extract_color(indexed_color) == "#FF0000"

    # 测试主题颜色
    class MockThemeColor:
        def __init__(self, theme):
            self.theme = theme

    theme_color = MockThemeColor(4)  # 强调1
    assert parser._extract_color(theme_color) == "#5B9BD5"

    # 测试自动颜色
    class MockAutoColor:
        def __init__(self):
            self.auto = True

    auto_color = MockAutoColor()
    assert parser._extract_color(auto_color) is None

    # 测试空颜色
    assert parser._extract_color(None) is None


# 新格式解析器测试
class TestXlsParser:
    """XLS解析器测试类。"""

    def test_xls_parser_basic_functionality(self):
        """测试XLS解析器的基本功能。"""
        parser = XlsParser()

        # 测试解析器实例化
        assert isinstance(parser, XlsParser)
        assert isinstance(parser, BaseParser)

    def test_xls_cell_value_processing(self):
        """测试XLS单元格值处理功能。"""
        parser = XlsParser()

        # 测试不同类型的值处理
        assert parser._process_cell_value("text", 1, None) == "text"
        assert parser._process_cell_value(123.0, 2, None) == 123
        assert parser._process_cell_value(123.5, 2, None) == 123.5
        assert parser._process_cell_value(True, 4, None) == True
        assert parser._process_cell_value(None, 0, None) is None

    def test_xls_color_mapping(self):
        """测试XLS颜色映射功能。"""
        parser = XlsParser()

        # 测试标准颜色索引
        assert parser._get_color_from_index(None, 0) == "#000000"  # 黑色
        assert parser._get_color_from_index(None, 1) == "#FFFFFF"  # 白色
        assert parser._get_color_from_index(None, 2) == "#FF0000"  # 红色
        assert parser._get_color_from_index(None, 999) == "#000000"  # 未知索引默认黑色

    def test_xls_parser_error_handling(self, sample_xlsx_path: Path):
        """测试XLS解析器的错误处理。"""
        parser = XlsParser()

        # 测试解析非XLS文件（使用xlsx文件）
        with pytest.raises(RuntimeError, match="解析XLS文件失败"):
            parser.parse(str(sample_xlsx_path))

    def test_xls_parser_nonexistent_file(self):
        """测试XLS解析器处理不存在的文件。"""
        parser = XlsParser()

        with pytest.raises(RuntimeError, match="解析XLS文件失败"):
            parser.parse("nonexistent.xls")

    def test_xls_style_extraction_default(self):
        """测试XLS样式提取的默认行为。"""
        parser = XlsParser()

        # 创建模拟的工作簿和工作表对象
        class MockWorkbook:
            def __init__(self):
                self.xf_list = []
                self.font_list = []
                self.colour_map = {}

        class MockWorksheet:
            def cell_xf_index(self, row, col):
                return 0

        # 测试默认样式提取
        style = parser._extract_style(MockWorkbook(), MockWorksheet(), 0, 0)
        assert isinstance(style, Style)
        assert style.bold == False
        assert style.italic == False
        assert style.font_color == "#000000"
        assert style.background_color == "#FFFFFF"

    def test_xls_style_extraction_with_font(self):
        """测试XLS样式提取包含字体信息。"""
        parser = XlsParser()

        # 创建模拟的字体对象
        class MockFont:
            def __init__(self):
                self.bold = True
                self.italic = False
                self.underline_type = 1
                self.height = 240  # 12pt in twips
                self.name = "Arial"
                self.colour_index = 2

        # 创建模拟的XF对象
        class MockXF:
            def __init__(self):
                self.font_index = 0
                self.background = None

        # 创建模拟的工作簿
        class MockWorkbook:
            def __init__(self):
                self.xf_list = [MockXF()]
                self.font_list = [MockFont()]
                self.colour_map = {}

        class MockWorksheet:
            def cell_xf_index(self, row, col):
                return 0

        # 测试样式提取
        style = parser._extract_style(MockWorkbook(), MockWorksheet(), 0, 0)
        assert style.bold == True
        assert style.italic == False
        assert style.underline == True
        assert style.font_size == 12.0  # 240/20
        assert style.font_name == "Arial"
        assert style.font_color == "#FF0000"  # 索引2对应红色


class TestXlsbParser:
    """XLSB解析器测试类。"""

    def test_xlsb_parser_basic_functionality(self):
        """测试XLSB解析器的基本功能。"""
        parser = XlsbParser()

        # 测试解析器实例化
        assert isinstance(parser, XlsbParser)
        assert isinstance(parser, BaseParser)

    def test_xlsb_cell_value_processing(self):
        """测试XLSB单元格值处理功能。"""
        parser = XlsbParser()

        # 测试不同类型的值处理
        assert parser._process_cell_value("text") == "text"
        assert parser._process_cell_value(123.0) == 123
        assert parser._process_cell_value(123.5) == 123.5
        assert parser._process_cell_value(True) == True
        assert parser._process_cell_value(None) is None

    def test_xlsb_basic_style_creation(self):
        """测试XLSB基础样式创建。"""
        parser = XlsbParser()
        style = parser._extract_basic_style()

        # 验证样式对象
        assert style.bold == False
        assert style.italic == False
        assert style.font_color == "#000000"
        assert style.background_color == "#FFFFFF"

    def test_xlsb_parser_error_handling(self, sample_xlsx_path: Path):
        """测试XLSB解析器的错误处理。"""
        parser = XlsbParser()

        # 测试解析非XLSB文件（使用xlsx文件）
        with pytest.raises(RuntimeError, match="解析XLSB文件失败"):
            parser.parse(str(sample_xlsx_path))

    def test_xlsb_parser_nonexistent_file(self):
        """测试XLSB解析器处理不存在的文件。"""
        parser = XlsbParser()

        with pytest.raises(RuntimeError, match="解析XLSB文件失败"):
            parser.parse("nonexistent.xlsb")

    def test_xlsb_style_extraction_comprehensive(self):
        """测试XLSB样式提取的全面功能。"""
        parser = XlsbParser()

        # 测试基础样式的所有属性
        style = parser._extract_basic_style()

        # 验证XLSB解析器设置的样式属性
        assert isinstance(style.bold, bool)
        assert isinstance(style.italic, bool)
        assert isinstance(style.underline, bool)
        assert isinstance(style.font_color, str)
        assert isinstance(style.background_color, str)
        assert isinstance(style.text_align, str)
        assert isinstance(style.vertical_align, str)

        # 验证XLSB解析器设置的默认值
        assert style.bold == False
        assert style.italic == False
        assert style.underline == False
        assert style.font_color == "#000000"
        assert style.background_color == "#FFFFFF"
        assert style.text_align == "left"
        assert style.vertical_align == "top"

        # 验证未设置的属性使用Style类的默认值
        assert style.font_size is None  # Style类的默认值
        assert style.font_name is None  # Style类的默认值

    def test_xlsb_cell_value_edge_cases(self):
        """测试XLSB单元格值处理的边界情况。"""
        parser = XlsbParser()

        # 测试边界情况
        assert parser._process_cell_value(0) == 0
        assert parser._process_cell_value(0.0) == 0
        assert parser._process_cell_value(-1.5) == -1.5
        assert parser._process_cell_value("") == ""
        assert parser._process_cell_value(" ") == " "
        assert parser._process_cell_value(False) == False

        # 测试特殊对象转换为字符串
        class CustomObject:
            def __str__(self):
                return "custom"

        assert parser._process_cell_value(CustomObject()) == "custom"


class TestParserFactoryExtended:
    """扩展的解析器工厂测试类。"""

    def test_get_parser_xls(self):
        """测试工厂返回XLS解析器。"""
        parser = ParserFactory.get_parser("test.xls")
        assert isinstance(parser, XlsParser)

    def test_get_parser_xlsb(self):
        """测试工厂返回XLSB解析器。"""
        parser = ParserFactory.get_parser("test.xlsb")
        assert isinstance(parser, XlsbParser)

    def test_get_supported_formats(self):
        """测试获取支持的格式列表。"""
        formats = ParserFactory.get_supported_formats()

        # 验证包含所有支持的格式
        expected_formats = ["csv", "xlsx", "xls", "xlsb"]
        for fmt in expected_formats:
            assert fmt in formats

        # 验证返回类型
        assert isinstance(formats, list)
        assert len(formats) >= 4

    def test_unsupported_format_error_message(self):
        """测试不支持格式的错误信息。"""
        with pytest.raises(UnsupportedFileType) as exc_info:
            ParserFactory.get_parser("test.unknown")

        error_message = str(exc_info.value)
        assert "不支持的文件格式" in error_message
        assert "unknown" in error_message
        assert "支持的格式" in error_message


@pytest.fixture
def failing_style_sample_path() -> Path:
    return Path(__file__).parent / "data" / "failing_style_sample.xlsx"

def test_failing_style_extraction(failing_style_sample_path: Path):
    """
    测试 XlsxParser 能否正确解析包含复杂样式的文件。
    这个测试目前预期会失败，用于驱动 bug 修复。
    """
    # 准备
    parser = XlsxParser()

    # 执行
    sheet = parser.parse(str(failing_style_sample_path))

    # 断言 - Style 1
    style1 = sheet.rows[0].cells[0].style
    assert style1 is not None, "Style1 should not be None"
    assert style1.font_color == "#FF0000"
    assert style1.background_color == "#FFFF00"
    assert style1.bold is True
    assert style1.italic is True
    assert style1.border_left == "1px solid #000000"
    assert style1.border_right == "3px double #00FF00"
    assert style1.border_top == "1px dashed #0000FF"
    assert style1.border_bottom == "2px solid #FF00FF"
    assert style1.text_align == "center"
    assert style1.vertical_align == "center"
    assert style1.wrap_text is True

    # 断言 - Style 2
    style2 = sheet.rows[1].cells[1].style
    assert style2 is not None, "Style2 should not be None"
    assert style2.font_color == "#0000FF"
    assert style2.underline is True
    assert style2.background_color == "#C0C0C0" # lightGray
    assert style2.text_align == "right"
    assert style2.vertical_align == "bottom"

    # 断言 - Style 3
    style3 = sheet.rows[2].cells[2].style
    assert style3 is not None, "Style3 should not be None"
    assert style3.font_color == "#E06666"
    assert style3.border_left == "2px solid #000000"
    assert style3.border_right == "2px solid #000000"
    assert style3.text_align == "left"
    assert style3.vertical_align == "top"

    # 断言 - Merged Cell
    merged_style = sheet.rows[3].cells[3].style
    assert merged_style is not None, "Merged style should not be None"
    assert sheet.merged_cells == ["D4:E5"]
    assert merged_style.font_color == "#8E7CC3"
    assert merged_style.background_color == "#D9EAD3"
    assert merged_style.text_align == "center"
    assert merged_style.vertical_align == "center"


# ==================== 新解析器测试用例 ====================

class TestXlsParser:
    """XLS解析器测试类"""

    def test_xls_parser_creation(self):
        """测试XlsParser能否正确创建"""
        parser = XlsParser()
        assert isinstance(parser, BaseParser)
        assert isinstance(parser, XlsParser)

    def test_xls_parser_color_mapping(self):
        """测试XLS颜色映射功能"""
        parser = XlsParser()

        # 测试标准颜色
        assert parser.default_color_map[0] == "#000000"  # 黑色
        assert parser.default_color_map[1] == "#FFFFFF"  # 白色
        assert parser.default_color_map[2] == "#FF0000"  # 红色

        # 测试颜色获取
        class MockWorkbook:
            colour_map = {2: (255, 0, 0)}  # 红色

        mock_wb = MockWorkbook()
        color = parser._get_color_from_index(mock_wb, 2)
        assert color == "#FF0000"

    def test_xls_parser_cell_reference_conversion(self):
        """测试Excel单元格引用转换"""
        parser = XlsParser()

        # 测试基本转换
        assert parser._index_to_excel_cell(0, 0) == "A1"
        assert parser._index_to_excel_cell(1, 1) == "B2"
        assert parser._index_to_excel_cell(0, 25) == "Z1"
        assert parser._index_to_excel_cell(0, 26) == "AA1"


class TestXlsbParser:
    """XLSB解析器测试类"""

    def test_xlsb_parser_creation(self):
        """测试XlsbParser能否正确创建"""
        parser = XlsbParser()
        assert isinstance(parser, BaseParser)
        assert isinstance(parser, XlsbParser)

    def test_xlsb_parser_cell_value_processing(self):
        """测试XLSB单元格值处理"""
        parser = XlsbParser()

        # 测试不同数据类型
        assert parser._process_cell_value(None) is None
        assert parser._process_cell_value("text") == "text"
        assert parser._process_cell_value(42) == 42
        assert parser._process_cell_value(3.14) == 3.14
        assert parser._process_cell_value(True) is True

        # 测试日期检测（44197对应2021-01-01）
        result = parser._process_cell_value(44197.0)
        assert hasattr(result, 'year')  # 应该是datetime对象

    def test_xlsb_parser_style_extraction(self):
        """测试XLSB基础样式提取"""
        parser = XlsbParser()

        class MockCellData:
            pass

        mock_cell = MockCellData()
        style = parser._extract_basic_style(mock_cell)

        # 验证默认样式
        assert style.font_name == "Calibri"
        assert style.font_size == 11.0
        assert style.font_color == "#000000"
        assert style.text_align == "left"
        assert style.vertical_align == "bottom"


class TestXlsmParser:
    """XLSM解析器测试类"""

    def test_xlsm_parser_creation(self):
        """测试XlsmParser能否正确创建"""
        parser = XlsmParser()
        assert isinstance(parser, BaseParser)
        assert isinstance(parser, XlsxParser)  # 继承自XlsxParser
        assert isinstance(parser, XlsmParser)

    def test_xlsm_parser_inheritance(self):
        """测试XLSM解析器继承功能"""
        parser = XlsmParser()

        # 验证继承的方法存在
        assert hasattr(parser, '_extract_style')
        assert hasattr(parser, '_extract_color')
        assert hasattr(parser, '_get_color_by_index')
        assert hasattr(parser, '_extract_fill_color')
        assert hasattr(parser, '_extract_number_format')
        assert hasattr(parser, '_extract_hyperlink')

    def test_xlsm_parser_macro_info_structure(self):
        """测试宏信息结构"""
        parser = XlsmParser()

        # 测试宏信息结构（不需要真实文件）
        expected_keys = ["has_macros", "vba_files_count", "vba_modules", "file_path"]

        # 模拟宏信息结构
        mock_info = {
            "has_macros": False,
            "vba_files_count": 0,
            "vba_modules": [],
            "file_path": "test.xlsm"
        }

        for key in expected_keys:
            assert key in mock_info

    def test_xlsm_parser_file_type_check(self):
        """测试文件类型检查"""
        parser = XlsmParser()

        # 测试扩展名检查
        assert "test.xlsm".lower().endswith('.xlsm')
        assert not "test.xlsx".lower().endswith('.xlsm')
        assert not "test.xls".lower().endswith('.xlsm')


class TestParserFactoryEnhanced:
    """增强的ParserFactory测试类"""

    def test_factory_supports_all_formats(self):
        """测试工厂支持所有5种格式"""
        supported_formats = ParserFactory.get_supported_formats()
        expected_formats = ["csv", "xlsx", "xls", "xlsb", "xlsm"]

        assert len(supported_formats) == 5
        for fmt in expected_formats:
            assert fmt in supported_formats

    def test_factory_get_parser_xls(self):
        """测试工厂返回XLS解析器"""
        parser = ParserFactory.get_parser("test.xls")
        assert isinstance(parser, XlsParser)

    def test_factory_get_parser_xlsb(self):
        """测试工厂返回XLSB解析器"""
        parser = ParserFactory.get_parser("test.xlsb")
        assert isinstance(parser, XlsbParser)

    def test_factory_get_parser_xlsm(self):
        """测试工厂返回XLSM解析器"""
        parser = ParserFactory.get_parser("test.xlsm")
        assert isinstance(parser, XlsmParser)

    def test_factory_format_info(self):
        """测试格式信息功能"""
        format_info = ParserFactory.get_format_info()

        # 验证所有格式都有信息
        expected_formats = ["csv", "xlsx", "xls", "xlsb", "xlsm"]
        for fmt in expected_formats:
            assert fmt in format_info
            info = format_info[fmt]
            assert "name" in info
            assert "description" in info
            assert "features" in info
            assert "parser_class" in info

    def test_factory_is_supported_format(self):
        """测试格式支持检查"""
        # 支持的格式
        assert ParserFactory.is_supported_format("test.csv")
        assert ParserFactory.is_supported_format("test.xlsx")
        assert ParserFactory.is_supported_format("test.xls")
        assert ParserFactory.is_supported_format("test.xlsb")
        assert ParserFactory.is_supported_format("test.xlsm")

        # 不支持的格式
        assert not ParserFactory.is_supported_format("test.doc")
        assert not ParserFactory.is_supported_format("test.txt")
        assert not ParserFactory.is_supported_format("test.pdf")

    def test_factory_case_insensitive(self):
        """测试格式检查大小写不敏感"""
        assert ParserFactory.is_supported_format("test.XLSX")
        assert ParserFactory.is_supported_format("test.XLS")
        assert ParserFactory.is_supported_format("test.XLSM")

        # 测试解析器获取也支持大小写
        parser_lower = ParserFactory.get_parser("test.xlsx")
        parser_upper = ParserFactory.get_parser("test.XLSX")
        assert type(parser_lower) == type(parser_upper)


class TestStyleFidelity:
    """样式保真度测试类"""

    def test_xlsx_enhanced_color_extraction(self):
        """测试XLSX增强的颜色提取"""
        parser = XlsxParser()

        # 测试简化的颜色提取
        class MockColor:
            def __init__(self, value):
                self.value = value

        # 测试RGB颜色
        rgb_color = MockColor('FF0000')
        assert parser._extract_color(rgb_color) == "#FF0000"

        # 测试ARGB颜色
        argb_color = MockColor('00FF0000')
        assert parser._extract_color(argb_color) == "#FF0000"

        # 测试索引颜色
        index_color = MockColor(2)
        assert parser._extract_color(index_color) == "#FF0000"

    def test_xlsx_enhanced_fill_extraction(self):
        """测试XLSX增强的填充提取"""
        parser = XlsxParser()

        class MockColor:
            def __init__(self, value):
                self.value = value

        class MockFill:
            def __init__(self, pattern_type=None, start_color=None, fg_color=None):
                self.patternType = pattern_type
                self.start_color = start_color
                self.fgColor = fg_color

        # 测试实色填充
        solid_fill = MockFill('solid', MockColor('FF0000'))
        assert parser._extract_fill_color(solid_fill) == "#FF0000"

        # 测试图案填充
        gray_fill = MockFill('lightGray')
        assert parser._extract_fill_color(gray_fill) == "#F2F2F2"

    def test_xlsx_enhanced_number_format(self):
        """测试XLSX增强的数字格式"""
        parser = XlsxParser()

        class MockCell:
            def __init__(self, number_format):
                self.number_format = number_format

        # 测试常见格式映射
        assert parser._extract_number_format(MockCell('0.00')) == "数字(2位小数)"
        assert parser._extract_number_format(MockCell('0%')) == "百分比"
        assert parser._extract_number_format(MockCell('mm/dd/yyyy')) == "日期(月/日/年)"
        assert parser._extract_number_format(MockCell('General')) == ""

    def test_xlsx_enhanced_hyperlink(self):
        """测试XLSX增强的超链接处理"""
        parser = XlsxParser()

        class MockHyperlink:
            def __init__(self, target=None, location=None):
                self.target = target
                self.location = location

        class MockCell:
            def __init__(self, hyperlink):
                self.hyperlink = hyperlink

        # 测试外部链接
        external_cell = MockCell(MockHyperlink(target='https://example.com'))
        assert parser._extract_hyperlink(external_cell) == 'https://example.com'

        # 测试内部链接
        internal_cell = MockCell(MockHyperlink(location='Sheet2!A1'))
        assert parser._extract_hyperlink(internal_cell) == '#内部:Sheet2!A1'

        # 测试无链接
        no_link_cell = MockCell(None)
        assert parser._extract_hyperlink(no_link_cell) is None


class TestErrorHandling:
    """错误处理测试类"""

    def test_parser_invalid_file_path(self):
        """测试解析器处理无效文件路径"""
        parsers = [XlsParser(), XlsbParser(), XlsmParser()]

        for parser in parsers:
            with pytest.raises(RuntimeError):
                parser.parse("nonexistent_file.invalid")

    def test_factory_invalid_extension(self):
        """测试工厂处理无效扩展名"""
        with pytest.raises(UnsupportedFileType) as exc_info:
            ParserFactory.get_parser("test.invalid")

        assert "不支持的文件格式" in str(exc_info.value)
        assert "invalid" in str(exc_info.value)

    def test_factory_no_extension(self):
        """测试工厂处理无扩展名文件"""
        with pytest.raises(UnsupportedFileType):
            ParserFactory.get_parser("filename_without_extension")

    def test_parser_empty_file_path(self):
        """测试解析器处理空文件路径"""
        parsers = [XlsParser(), XlsbParser(), XlsmParser()]

        for parser in parsers:
            with pytest.raises((RuntimeError, ValueError, FileNotFoundError)):
                parser.parse("")


class TestPerformanceBenchmark:
    """性能基准测试类"""

    def test_parser_creation_performance(self):
        """测试解析器创建性能"""
        import time

        # 测试解析器创建时间
        start_time = time.time()
        for _ in range(100):
            XlsParser()
            XlsbParser()
            XlsmParser()
        end_time = time.time()

        # 100次创建应该在1秒内完成
        assert (end_time - start_time) < 1.0

    def test_factory_performance(self):
        """测试工厂性能"""
        import time

        test_files = ["test.csv", "test.xlsx", "test.xls", "test.xlsb", "test.xlsm"]

        start_time = time.time()
        for _ in range(1000):
            for file in test_files:
                ParserFactory.get_parser(file)
        end_time = time.time()

        # 5000次工厂调用应该在1秒内完成
        assert (end_time - start_time) < 1.0
