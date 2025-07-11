"""
集成测试模块

测试端到端的工作流程，验证各组件之间的集成。
"""

import pytest
import tempfile
import os
from pathlib import Path
from src.parsers.factory import ParserFactory
from src.converters.html_converter import HTMLConverter
from src.converters.json_converter import JSONConverter
from src.services.sheet_service import SheetService
from src.mcp_server.tools import (
    _handle_parse_sheet_to_json,
    _handle_convert_json_to_html,
    _handle_convert_file_to_html,
    _handle_get_table_summary
)
from src.utils.performance import PerformanceOptimizer


class TestEndToEndWorkflows:
    """端到端工作流程测试。"""
    
    def test_csv_to_html_workflow(self):
        """测试CSV到HTML的完整工作流程。"""
        # 1. 解析CSV文件
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser("tests/data/sample.csv")
        sheet = parser.parse("tests/data/sample.csv")
        
        # 2. 转换为HTML
        html_converter = HTMLConverter()
        html_content = html_converter.convert(sheet)
        
        # 3. 验证结果
        assert isinstance(html_content, str)
        assert "<table>" in html_content
        assert "</table>" in html_content
        assert len(html_content) > 100  # 确保有实际内容
    
    def test_xlsx_to_html_workflow(self):
        """测试XLSX到HTML的完整工作流程。"""
        # 1. 解析XLSX文件
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser("tests/data/sample.xlsx")
        sheet = parser.parse("tests/data/sample.xlsx")
        
        # 2. 转换为HTML
        html_converter = HTMLConverter()
        html_content = html_converter.convert(sheet)
        
        # 3. 验证结果
        assert isinstance(html_content, str)
        assert "<table>" in html_content
        assert "ID" in html_content
        assert "Name" in html_content
        assert "Value" in html_content
    
    def test_json_roundtrip_workflow(self):
        """测试JSON往返转换工作流程。"""
        # 1. 解析文件到Sheet
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser("tests/data/sample.xlsx")
        original_sheet = parser.parse("tests/data/sample.xlsx")
        
        # 2. 转换为JSON
        json_converter = JSONConverter()
        json_data = json_converter.convert(original_sheet)
        
        # 3. 从JSON重建Sheet
        from src.mcp_server.tools import _json_to_sheet
        reconstructed_sheet = _json_to_sheet(json_data)
        
        # 4. 验证数据完整性
        assert reconstructed_sheet.name == original_sheet.name
        assert len(reconstructed_sheet.rows) == len(original_sheet.rows)
        
        # 验证第一行数据
        original_first_row = [cell.value for cell in original_sheet.rows[0].cells]
        reconstructed_first_row = [cell.value for cell in reconstructed_sheet.rows[0].cells]
        assert original_first_row == reconstructed_first_row
    
    def test_performance_optimization_workflow(self):
        """测试性能优化工作流程。"""
        # 1. 创建性能优化器
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()
        
        # 2. 解析文件
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser("tests/data/sample.xlsx")
        sheet = parser.parse("tests/data/sample.xlsx")
        
        # 3. 分析性能
        metrics = optimizer.analyze_sheet_performance(sheet)
        
        # 4. 验证性能指标
        assert metrics.rows > 0
        assert metrics.cols > 0
        assert metrics.total_cells > 0
        assert metrics.estimated_html_size > 0
        assert metrics.estimated_json_size > 0
        assert metrics.processing_time >= 0
        assert metrics.recommendation in [
            "small_file_direct", "medium_file_direct", 
            "medium_file_with_warning", "large_file_recommend_file_output",
            "very_large_file_require_pagination"
        ]
    
    def test_service_layer_integration(self):
        """测试服务层集成。"""
        # 1. 创建服务
        parser_factory = ParserFactory()
        html_converter = HTMLConverter()
        service = SheetService(parser_factory, html_converter)
        
        # 2. 测试CSV处理
        csv_html = service.convert_to_html("tests/data/sample.csv")
        assert "<table>" in csv_html
        assert len(csv_html) > 100
        
        # 3. 测试XLSX处理
        xlsx_html = service.convert_to_html("tests/data/sample.xlsx")
        assert "<table>" in xlsx_html
        assert len(xlsx_html) > 100


class TestMCPToolsIntegration:
    """MCP工具集成测试。"""
    
    @pytest.mark.asyncio
    async def test_parse_sheet_to_json_integration(self):
        """测试parse_sheet_to_json工具的集成。"""
        # 创建模拟服务
        parser_factory = ParserFactory()
        html_converter = HTMLConverter()
        service = SheetService(parser_factory, html_converter)
        
        # 测试工具调用
        arguments = {"file_path": "tests/data/sample.xlsx"}
        result = await _handle_parse_sheet_to_json(arguments, service)
        
        # 验证结果
        assert len(result) == 1
        assert "JSON 转换成功" in result[0].text
        assert "行数:" in result[0].text
        assert "列数:" in result[0].text
    
    @pytest.mark.asyncio
    async def test_convert_file_to_html_integration(self):
        """测试convert_file_to_html工具的集成。"""
        # 创建模拟服务
        parser_factory = ParserFactory()
        html_converter = HTMLConverter()
        service = SheetService(parser_factory, html_converter)
        
        # 测试工具调用
        arguments = {"file_path": "tests/data/sample.csv"}
        result = await _handle_convert_file_to_html(arguments, service)
        
        # 验证结果
        assert len(result) == 1
        assert "HTML 转换完成" in result[0].text
        assert "<table>" in result[0].text
    
    @pytest.mark.asyncio
    async def test_get_table_summary_integration(self):
        """测试get_table_summary工具的集成。"""
        # 创建模拟服务
        parser_factory = ParserFactory()
        html_converter = HTMLConverter()
        service = SheetService(parser_factory, html_converter)
        
        # 测试工具调用
        arguments = {"file_path": "tests/data/sample.xlsx"}
        result = await _handle_get_table_summary(arguments, service)
        
        # 验证结果
        assert len(result) == 1
        assert "表格摘要:" in result[0].text
        assert "基本统计:" in result[0].text
        assert "性能指标:" in result[0].text
        assert "处理建议:" in result[0].text
    
    @pytest.mark.asyncio
    async def test_json_to_html_file_integration(self):
        """测试JSON到HTML文件的集成。"""
        # 1. 先获取JSON数据
        parser_factory = ParserFactory()
        html_converter = HTMLConverter()
        service = SheetService(parser_factory, html_converter)
        
        arguments = {"file_path": "tests/data/sample.xlsx"}
        json_result = await _handle_parse_sheet_to_json(arguments, service)
        
        # 2. 从结果中提取JSON数据（简化处理）
        import json
        json_text = json_result[0].text
        json_start = json_text.find('{"metadata"')
        if json_start == -1:
            json_start = json_text.find('{"data"')
        
        if json_start != -1:
            json_data_str = json_text[json_start:]
            json_data = json.loads(json_data_str)
            
            # 3. 转换为HTML文件
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as tmp_file:
                tmp_path = tmp_file.name
            
            try:
                arguments = {
                    "json_data": json_data,
                    "output_path": tmp_path
                }
                result = await _handle_convert_json_to_html(arguments, service)
                
                # 4. 验证结果
                assert len(result) == 1
                assert "HTML 文件生成成功" in result[0].text
                assert os.path.exists(tmp_path)
                
                # 验证文件内容
                with open(tmp_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                assert "<table>" in html_content
                
            finally:
                # 清理临时文件
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)


class TestErrorHandling:
    """错误处理集成测试。"""
    
    def test_unsupported_file_format(self):
        """测试不支持的文件格式处理。"""
        parser_factory = ParserFactory()
        
        with pytest.raises(Exception):  # 应该抛出UnsupportedFileType异常
            parser_factory.get_parser("test.unknown")
    
    def test_file_not_found_handling(self):
        """测试文件不存在的处理。"""
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser("nonexistent.csv")
        
        with pytest.raises(Exception):  # 应该抛出文件不存在异常
            parser.parse("nonexistent.csv")
    
    @pytest.mark.asyncio
    async def test_mcp_tool_error_handling(self):
        """测试MCP工具的错误处理。"""
        parser_factory = ParserFactory()
        html_converter = HTMLConverter()
        service = SheetService(parser_factory, html_converter)
        
        # 测试文件不存在
        arguments = {"file_path": "nonexistent.xlsx"}
        
        with pytest.raises(Exception):
            await _handle_parse_sheet_to_json(arguments, service)
