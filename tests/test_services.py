import pytest
from src.services.sheet_service import SheetService
from src.parsers.factory import ParserFactory
from src.converters.html_converter import HTMLConverter

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
    assert "<td>Name</td>" in html_output
    assert "<td>Alice</td>" in html_output
    assert "<td>30</td>" in html_output

def test_process_file_xlsx(sheet_service):
    """
    测试 XLSX 文件到 HTML 的完整处理流程。
    """
    file_path = 'tests/data/sample.xlsx'
    html_output = sheet_service.convert_to_html(file_path)
    assert "<td>Name</td>" in html_output
    assert "<td>Bob</td>" in html_output
    assert "<td>25</td>" in html_output