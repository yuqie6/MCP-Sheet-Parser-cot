"""
服务层测试模块

测试业务逻辑层的功能。
"""

import pytest
from unittest.mock import Mock
from src.services.sheet_service import SheetService
from src.parsers.factory import ParserFactory
from src.converters.html_converter import HTMLConverter
from src.models.table_model import Sheet, Row, Cell, Style

@pytest.fixture
def sheet_service():
    """为测试提供 SheetService 实例。"""
    parser_factory = ParserFactory()
    html_converter = HTMLConverter()
    return SheetService(parser_factory=parser_factory, html_converter=html_converter)

def test_sheet_service_initialization(sheet_service):
    """
    测试 SheetService 能否用其依赖项初始化。
    """
    assert sheet_service is not None

def test_process_file_csv(sheet_service):
    """
    测试 CSV 文件到 HTML 的完整处理流程。
    """
    file_path = 'tests/data/sample.csv'
    html_output = sheet_service.convert_to_html(file_path)
    # 检查实际的CSV数据内容
    assert "header" in html_output.lower() or "value" in html_output.lower()
    assert "<td>" in html_output
    assert "</td>" in html_output

def test_process_file_xlsx(sheet_service):
    """
    测试 XLSX 文件到 HTML 的完整处理流程。
    """
    file_path = 'tests/data/sample.xlsx'
    html_output = sheet_service.convert_to_html(file_path)
    # 检查实际的XLSX数据内容
    assert "Name" in html_output or "ID" in html_output
    assert "Alice" in html_output or "Bob" in html_output
    assert "<td" in html_output  # 检查td标签存在（可能带有class属性）


def test_service_with_mock_components():
    """测试服务与模拟组件的交互。"""
    # 创建模拟对象
    parser_factory = Mock()
    html_converter = Mock()
    parser = Mock()

    # 创建测试数据
    test_sheet = Sheet(
        name="test",
        rows=[
            Row(cells=[Cell(value="A1"), Cell(value="B1")]),
            Row(cells=[Cell(value="A2"), Cell(value="B2")])
        ]
    )

    # 配置模拟行为
    parser_factory.get_parser.return_value = parser
    parser.parse.return_value = test_sheet
    html_converter.convert.return_value = "<table><tr><td>A1</td><td>B1</td></tr></table>"

    # 创建服务并测试
    service = SheetService(parser_factory, html_converter)
    result = service.convert_to_html("test.csv")

    # 验证调用（服务层传递文件扩展名）
    parser_factory.get_parser.assert_called_once_with("csv")
    parser.parse.assert_called_once_with("test.csv")
    html_converter.convert.assert_called_once_with(test_sheet)

    # 验证结果
    assert result == "<table><tr><td>A1</td><td>B1</td></tr></table>"


def test_service_error_handling():
    """测试服务的错误处理。"""
    # 创建模拟对象
    parser_factory = Mock()
    html_converter = Mock()
    parser = Mock()

    # 配置模拟行为：解析器抛出异常
    parser_factory.get_parser.return_value = parser
    parser.parse.side_effect = Exception("File not found")

    # 创建服务并测试
    service = SheetService(parser_factory, html_converter)

    with pytest.raises(Exception, match="File not found"):
        service.convert_to_html("nonexistent.xlsx")


def test_service_file_extension_handling():
    """测试不同文件扩展名的处理。"""
    # 创建模拟对象
    parser_factory = Mock()
    html_converter = Mock()
    parser = Mock()

    # 创建测试数据
    test_sheet = Sheet(name="test", rows=[])

    # 配置模拟行为
    parser_factory.get_parser.return_value = parser
    parser.parse.return_value = test_sheet
    html_converter.convert.return_value = "<table></table>"

    # 创建服务
    service = SheetService(parser_factory, html_converter)

    # 测试不同扩展名
    test_cases = [
        ("test.csv", "csv"),
        ("test.xlsx", "xlsx"),
        ("test.xls", "xls"),
        ("test.xlsb", "xlsb")
    ]

    for file_path, expected_ext in test_cases:
        service.convert_to_html(file_path)
        parser_factory.get_parser.assert_called_with(expected_ext)