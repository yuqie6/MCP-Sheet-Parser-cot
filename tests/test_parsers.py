import pytest
from pathlib import Path
from src.models.table_model import Sheet, Row, Cell
from src.parsers.base_parser import BaseParser
from src.parsers.csv_parser import CsvParser
from src.parsers.xlsx_parser import XlsxParser
from src.parsers.xls_parser import XlsParser
from src.parsers.xlsb_parser import XlsbParser
from src.parsers.factory import ParserFactory, UnsupportedFileType

@pytest.fixture
def sample_csv_path() -> Path:
    return Path(__file__).parent / "data" / "sample.csv"

@pytest.fixture
def sample_xlsx_path() -> Path:
    return Path(__file__).parent / "data" / "sample.xlsx"

def test_csv_parser_success(sample_csv_path: Path):
    """
    Tests that the CsvParser can successfully parse a simple CSV file.
    """
    # Arrange
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

    # Act
    result_sheet = parser.parse(str(sample_csv_path))

    # Assert
    assert isinstance(result_sheet, Sheet)
    assert result_sheet.name == "sample"
    assert len(result_sheet.rows) == 4
    
    # Check headers
    assert [cell.value for cell in result_sheet.rows[0].cells] == ["header1", "header2", "header3"]
    
    # Check data rows
    assert [cell.value for cell in result_sheet.rows[1].cells] == ["value1_1", "value1_2", "value1_3"]
    assert [cell.value for cell in result_sheet.rows[2].cells] == ["value2_1", "value2_2", "value2_3"]
    assert [cell.value for cell in result_sheet.rows[3].cells] == ["value3_1", "value3_2", "value3_3"]

    # Deep comparison
    assert result_sheet == expected_sheet

def test_xlsx_parser_success(sample_xlsx_path: Path):
    """
    Tests that the XlsxParser can successfully parse a simple XLSX file.
    """
    # Arrange
    parser = XlsxParser()
    expected_sheet = Sheet(
        name="Sheet1",
        rows=[
            Row(cells=[Cell(value="ID"), Cell(value="Name"), Cell(value="Value")]),
            Row(cells=[Cell(value=1), Cell(value="Alice"), Cell(value=100)]),
            Row(cells=[Cell(value=2), Cell(value="Bob"), Cell(value=200)]),
        ]
    )

    # Act
    result_sheet = parser.parse(str(sample_xlsx_path))

    # Assert
    assert isinstance(result_sheet, Sheet)
    assert result_sheet.name == "Sheet1"
    assert len(result_sheet.rows) == 3
    
    # Check headers
    assert [cell.value for cell in result_sheet.rows[0].cells] == ["ID", "Name", "Value"]
    
    # Check data rows
    assert [cell.value for cell in result_sheet.rows[1].cells] == [1, "Alice", 100]
    assert [cell.value for cell in result_sheet.rows[2].cells] == [2, "Bob", 200]

    # Deep comparison
    assert result_sheet == expected_sheet


class TestParserFactory:
    def test_get_parser_csv(self, sample_csv_path: Path):
        """
        Tests that the ParserFactory returns a CsvParser instance for .csv files.
        """
        # Act
        parser = ParserFactory.get_parser(str(sample_csv_path))
        # Assert
        assert isinstance(parser, CsvParser)

    def test_get_parser_xlsx(self, sample_xlsx_path: Path):
        """
        Tests that the ParserFactory returns an XlsxParser instance for .xlsx files.
        """
        # Act
        parser = ParserFactory.get_parser(str(sample_xlsx_path))
        # Assert
        assert isinstance(parser, XlsxParser)

    def test_get_parser_unsupported_file_type(self):
        """
        Tests that the ParserFactory raises an UnsupportedFileType error for unsupported file types.
        """
        # Arrange
        unsupported_file_path = "test.txt"
        # Act & Assert
        with pytest.raises(UnsupportedFileType):
            ParserFactory.get_parser(unsupported_file_path)


def test_xlsx_style_extraction_basic_properties(sample_xlsx_path: Path):
    """Test basic style extraction from XLSX files."""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    # Test that styles are extracted
    first_cell = sheet.rows[0].cells[0]
    assert first_cell.style is not None

    # Test font properties
    style = first_cell.style
    assert isinstance(style.bold, bool)
    assert isinstance(style.italic, bool)
    assert isinstance(style.underline, bool)
    assert isinstance(style.font_color, str)
    assert style.font_color.startswith('#')

    # Test font size and name
    if style.font_size is not None:
        assert isinstance(style.font_size, (int, float))
        assert style.font_size > 0

    if style.font_name is not None:
        assert isinstance(style.font_name, str)
        assert len(style.font_name) > 0


def test_xlsx_style_extraction_background_and_alignment(sample_xlsx_path: Path):
    """Test background color and alignment extraction."""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    first_cell = sheet.rows[0].cells[0]
    style = first_cell.style

    # Test background color
    assert isinstance(style.background_color, str)
    assert style.background_color.startswith('#')

    # Test alignment properties
    assert style.text_align in ['left', 'center', 'right', 'justify', 'general']
    assert style.vertical_align in ['top', 'middle', 'bottom', 'center']

    # Test wrap text
    assert isinstance(style.wrap_text, bool)


def test_xlsx_style_extraction_borders(sample_xlsx_path: Path):
    """Test border style extraction."""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    first_cell = sheet.rows[0].cells[0]
    style = first_cell.style

    # Test border properties
    assert isinstance(style.border_top, str)
    assert isinstance(style.border_bottom, str)
    assert isinstance(style.border_left, str)
    assert isinstance(style.border_right, str)
    assert isinstance(style.border_color, str)
    assert style.border_color.startswith('#')


def test_xlsx_style_extraction_number_format(sample_xlsx_path: Path):
    """Test number format extraction."""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    first_cell = sheet.rows[0].cells[0]
    style = first_cell.style

    # Test number format
    assert isinstance(style.number_format, str)


def test_style_fidelity_coverage(sample_xlsx_path: Path):
    """Test that style extraction covers all major style properties for 95% fidelity."""
    parser = XlsxParser()
    sheet = parser.parse(str(sample_xlsx_path))

    # Test multiple cells to ensure comprehensive coverage
    for row_idx, row in enumerate(sheet.rows[:3]):  # Test first 3 rows
        for cell_idx, cell in enumerate(row.cells[:3]):  # Test first 3 columns
            style = cell.style

            # Verify all style properties are present and have correct types
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

            # Test color format validity
            assert style.font_color.startswith('#'), f"Invalid font color format: {style.font_color}"
            assert style.background_color.startswith('#'), f"Invalid background color format: {style.background_color}"
            assert style.border_color.startswith('#'), f"Invalid border color format: {style.border_color}"


def test_indexed_color_mapping():
    """Test indexed color mapping functionality."""
    parser = XlsxParser()

    # Test indexed color conversion
    test_colors = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    for index in test_colors:
        color = parser._get_indexed_color(index)
        assert isinstance(color, str)
        assert color.startswith('#')
        assert len(color) == 7  # #RRGGBB format


def test_border_style_mapping():
    """Test border style mapping functionality."""
    parser = XlsxParser()

    # Create mock border side objects for testing
    class MockBorderSide:
        def __init__(self, style):
            self.style = style

    # Test various border styles
    border_styles = ['thin', 'medium', 'thick', 'double', 'dotted', 'dashed', 'hair']
    for style_name in border_styles:
        mock_border = MockBorderSide(style_name)
        css_style = parser._get_border_style(mock_border)
        assert isinstance(css_style, str)
        assert len(css_style) > 0
        assert 'px' in css_style  # Should contain pixel values

    # Test empty border
    empty_border = MockBorderSide(None)
    css_style = parser._get_border_style(empty_border)
    assert css_style == ""


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
        assert style.text_align == "left"


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
