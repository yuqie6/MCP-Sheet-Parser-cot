import pytest
from pathlib import Path
from src.converters.paginated_html_converter import PaginatedHTMLConverter
from src.models.table_model import Sheet, Row, Cell

@pytest.fixture
def sample_sheet():
    """创建一个包含15行数据的Sheet对象用于测试。"""
    rows = [Row(cells=[Cell(value=f'R{i}C1'), Cell(value=f'R{i}C2')]) for i in range(1, 16)]
    return Sheet(name='TestSheet', rows=rows)

@pytest.fixture
def large_sheet():
    """创建一个包含100行数据的Sheet对象用于测试。"""
    rows = [Row(cells=[Cell(value=f'R{i}C1')]) for i in range(1, 101)]
    return Sheet(name='LargeTestSheet', rows=rows)

def test_initialization():
    """测试PaginatedHTMLConverter的初始化。"""
    converter = PaginatedHTMLConverter(page_size=50, page_number=2)
    assert converter.page_size == 50
    assert converter.page_number == 2

def test_initialization_invalid_values():
    """测试初始化时传入无效值（小于1）的情况。"""
    converter = PaginatedHTMLConverter(page_size=0, page_number=-1)
    assert converter.page_size == 1
    assert converter.page_number == 1

def test_create_paginated_sheet_page_1(sample_sheet):
    """测试第一页的分页Sheet创建。"""
    converter = PaginatedHTMLConverter(page_size=5, page_number=1)
    paginated_sheet = converter._create_paginated_sheet(sample_sheet, 0, 5)
    assert len(paginated_sheet.rows) == 5
    assert paginated_sheet.rows[0].cells[0].value == 'R1C1'
    assert paginated_sheet.rows[4].cells[0].value == 'R5C1'

def test_create_paginated_sheet_page_2(sample_sheet):
    """测试第二页的分页Sheet创建。"""
    converter = PaginatedHTMLConverter(page_size=5, page_number=2)
    paginated_sheet = converter._create_paginated_sheet(sample_sheet, 5, 10)
    assert len(paginated_sheet.rows) == 5
    assert paginated_sheet.rows[0].cells[0].value == 'R6C1'
    assert paginated_sheet.rows[4].cells[0].value == 'R10C1'

def test_create_paginated_sheet_last_page(sample_sheet):
    """测试最后一页（不完整）的分页Sheet创建。"""
    converter = PaginatedHTMLConverter(page_size=5, page_number=3)
    paginated_sheet = converter._create_paginated_sheet(sample_sheet, 10, 15)
    assert len(paginated_sheet.rows) == 5
    assert paginated_sheet.rows[0].cells[0].value == 'R11C1'
    assert paginated_sheet.rows[4].cells[0].value == 'R15C1'

def test_generate_html_pagination_info(sample_sheet):
    """测试生成的HTML是否包含正确的分页信息。"""
    converter = PaginatedHTMLConverter(page_size=10, page_number=2)
    html = converter._generate_html(sample_sheet)
    assert '显示第 11-15 行，共 15 行' in html
    assert '第 2 页，共 2 页' in html

def test_convert_to_file_success(sample_sheet, tmp_path):
    """测试convert_to_file成功写入文件。"""
    converter = PaginatedHTMLConverter()
    output_file = tmp_path / "output.html"
    result = converter.convert_to_file(sample_sheet, str(output_file))
    
    assert result["status"] == "success"
    assert Path(result["output_path"]).exists()
    assert result["file_size"] > 0

def test_convert_to_file_error(sample_sheet, monkeypatch):
    """测试convert_to_file在写入失败时返回错误。"""
    def mock_open_raises_exception(*args, **kwargs):
        raise IOError("Permission denied")

    monkeypatch.setattr("builtins.open", mock_open_raises_exception)
    
    converter = PaginatedHTMLConverter()
    result = converter.convert_to_file(sample_sheet, "/non_existent_dir/output.html")
    
    assert result["status"] == "error"
    assert "Permission denied" in result["error"]

def test_pagination_controls_first_page(large_sheet):
    """测试第一页时分页控件的显示。"""
    converter = PaginatedHTMLConverter(page_size=10, page_number=1)
    controls = converter._generate_pagination_controls(1, 10, 100, 0, 10)
    assert '<button class="page-btn disabled" disabled>上一页</button>' in controls
    assert '<button onclick="loadPage(2)" class="page-btn">下一页</button>' in controls

def test_pagination_controls_last_page(large_sheet):
    """测试最后一页时分页控件的显示。"""
    converter = PaginatedHTMLConverter(page_size=10, page_number=10)
    controls = converter._generate_pagination_controls(10, 10, 100, 90, 100)
    assert '<button onclick="loadPage(9)" class="page-btn">上一页</button>' in controls
    assert '<button class="page-btn disabled" disabled>下一页</button>' in controls

def test_pagination_controls_middle_page_with_ellipsis(large_sheet):
    """测试中间页码时，省略号是否正确显示。"""
    converter = PaginatedHTMLConverter(page_size=10, page_number=5)
    controls = converter._generate_pagination_controls(5, 10, 100, 40, 50)
    assert '<span class="page-ellipsis">...</span>' in controls
    assert '<button onclick="loadPage(1)" class="page-btn">1</button>' in controls
    assert '<button onclick="loadPage(10)" class="page-btn">10</button>' in controls
