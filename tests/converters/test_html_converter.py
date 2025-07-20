"""
HTML转换器的TDD测试套件。

本测试套件专注于测试HTMLConverter的核心功能，
包括文件转换、错误处理、表头检测等功能。
"""

import pytest
from unittest.mock import MagicMock, patch, mock_open
from pathlib import Path
import tempfile
import os

from src.converters.html_converter import HTMLConverter
from src.models.table_model import Sheet, Row, Cell, Style


@pytest.fixture
def sample_sheet():
    """创建测试用的Sheet对象"""
    # 创建样式
    header_style = Style(
        bold=True,
        background_color="#E0E0E0",
        text_align="center"
    )

    data_style = Style(
        bold=False,
        background_color="#FFFFFF"
    )
    
    # 创建表头行
    header_row = Row(cells=[
        Cell(value="姓名", style=header_style),
        Cell(value="年龄", style=header_style),
        Cell(value="部门", style=header_style)
    ])
    
    # 创建数据行
    data_row1 = Row(cells=[
        Cell(value="张三", style=data_style),
        Cell(value=25, style=data_style),
        Cell(value="技术部", style=data_style)
    ])
    
    data_row2 = Row(cells=[
        Cell(value="李四", style=data_style),
        Cell(value=30, style=data_style),
        Cell(value="销售部", style=data_style)
    ])
    
    return Sheet(
        name="员工信息",
        rows=[header_row, data_row1, data_row2],
        merged_cells=[]
    )


def test_html_converter_init():
    """
    TDD测试：HTMLConverter应该正确初始化
    """
    converter = HTMLConverter(compact_mode=True, header_rows=2, auto_detect_headers=False)
    
    # 验证初始化参数
    assert converter.compact_mode is True
    assert converter.header_rows == 2
    assert converter.auto_detect_headers is False
    
    # 验证子转换器被正确初始化
    assert converter.style_converter is not None
    assert converter.cell_converter is not None
    assert converter.table_converter is not None
    assert converter.chart_converter is not None


def test_convert_to_files_single_sheet(sample_sheet):
    """
    TDD测试：convert_to_files应该处理单个工作表
    
    这个测试覆盖第51-52行的单工作表路径
    """
    converter = HTMLConverter()
    
    with tempfile.TemporaryDirectory() as temp_dir:
        output_path = os.path.join(temp_dir, "test.html")
        
        # 模拟_generate_html方法
        with patch.object(converter, '_generate_html', return_value="<html>test</html>"):
            results = converter.convert_to_files([sample_sheet], output_path)
        
        # 验证结果
        assert len(results) == 1
        result = results[0]
        assert result["status"] == "success"
        assert result["sheet_name"] == "员工信息"
        assert result["rows_converted"] == 3
        assert result["cells_converted"] == 9
        assert "output_path" in result
        assert "file_size" in result


def test_convert_to_files_multiple_sheets():
    """
    TDD测试：convert_to_files应该处理多个工作表
    
    这个测试覆盖第49-50行的多工作表路径
    """
    converter = HTMLConverter()
    
    # 创建两个工作表
    sheet1 = Sheet(name="工作表1", rows=[], merged_cells=[])
    sheet2 = Sheet(name="工作表2", rows=[], merged_cells=[])
    
    with tempfile.TemporaryDirectory() as temp_dir:
        output_path = os.path.join(temp_dir, "test.html")
        
        # 模拟_generate_html方法
        with patch.object(converter, '_generate_html', return_value="<html>test</html>"):
            results = converter.convert_to_files([sheet1, sheet2], output_path)
        
        # 验证结果
        assert len(results) == 2
        
        # 验证文件名包含工作表名称
        assert "工作表1" in results[0]["output_path"]
        assert "工作表2" in results[1]["output_path"]


def test_convert_to_files_with_exception(sample_sheet):
    """
    TDD测试：convert_to_files应该处理转换异常
    
    这个测试覆盖第70-72行的异常处理代码
    """
    converter = HTMLConverter()
    
    with tempfile.TemporaryDirectory() as temp_dir:
        output_path = os.path.join(temp_dir, "test.html")
        
        # 模拟_generate_html抛出异常
        with patch.object(converter, '_generate_html', side_effect=Exception("转换失败")):
            results = converter.convert_to_files([sample_sheet], output_path)
        
        # 验证错误结果
        assert len(results) == 1
        result = results[0]
        assert result["status"] == "error"
        assert result["sheet_name"] == "员工信息"
        assert "转换失败" in result["error"]


def test_generate_html_with_compact_mode(sample_sheet):
    """
    TDD测试：_generate_html应该支持紧凑模式

    这个测试覆盖第109-110行的紧凑模式代码
    """
    converter = HTMLConverter(compact_mode=True)

    # 模拟compact_html函数
    with patch('src.converters.html_converter.compact_html') as mock_compact:
        mock_compact.return_value = "compacted_html"

        result = converter._generate_html(sample_sheet)

        # 验证compact_html被调用
        mock_compact.assert_called_once()
        assert result == "compacted_html"


def test_detect_header_rows_empty_sheet():
    """
    TDD测试：_detect_header_rows应该处理空工作表
    
    这个测试覆盖第119-120行的边界情况
    """
    converter = HTMLConverter()
    
    # 测试空工作表
    empty_sheet = Sheet(name="空表", rows=[], merged_cells=[])
    assert converter._detect_header_rows(empty_sheet) == 0
    
    # 测试只有一行的工作表
    single_row_sheet = Sheet(
        name="单行表", 
        rows=[Row(cells=[Cell(value="标题")])], 
        merged_cells=[]
    )
    assert converter._detect_header_rows(single_row_sheet) == 1


def test_detect_header_rows_with_merged_cells():
    """
    TDD测试：_detect_header_rows应该处理合并单元格
    
    这个测试覆盖第126-140行的合并单元格检测代码
    """
    converter = HTMLConverter()
    
    # 创建有合并单元格的工作表
    sheet = Sheet(
        name="合并表",
        rows=[
            Row(cells=[Cell(value="标题1"), Cell(value="标题2")]),
            Row(cells=[Cell(value="子标题1"), Cell(value="子标题2")]),
            Row(cells=[Cell(value="数据1"), Cell(value="数据2")])
        ],
        merged_cells=["A1:A2"]  # 跨两行的合并单元格
    )
    
    # 模拟range_parser
    with patch('src.utils.range_parser.parse_range_string', return_value=(0, 0, 1, 0)):
        header_rows = converter._detect_header_rows(sheet)

        # 应该检测到2行表头（因为合并单元格跨越到第2行）
        assert header_rows == 2


def test_detect_header_rows_with_parse_exception():
    """
    TDD测试：_detect_header_rows应该处理解析异常
    
    这个测试覆盖第134-135行的异常处理代码
    """
    converter = HTMLConverter()
    
    sheet = Sheet(
        name="异常表",
        rows=[
            Row(cells=[Cell(value="标题")]),
            Row(cells=[Cell(value="数据")])
        ],
        merged_cells=["INVALID_RANGE"]  # 无效的合并单元格范围
    )
    
    # 模拟parse_range_string抛出异常
    with patch('src.utils.range_parser.parse_range_string', side_effect=Exception("解析错误")):
        header_rows = converter._detect_header_rows(sheet)

        # 应该返回默认的1行表头
        assert header_rows == 1
