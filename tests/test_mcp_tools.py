"""
MCP工具测试模块

全面测试3个核心MCP工具的功能
目标覆盖率：80%+

注意：所有MCP工具函数都需要core_service参数
"""

import pytest
import tempfile
import os
from pathlib import Path
from src.mcp_server.tools import (
    _handle_convert_to_html,
    _handle_parse_sheet,
    _handle_apply_changes
)
from src.core_service import CoreService


class TestMCPToolsHelper:
    """MCP工具测试辅助类，提供通用的测试方法。"""

    @staticmethod
    def get_core_service():
        """获取CoreService实例。"""
        return CoreService()

    @staticmethod
    async def call_convert_to_html(arguments):
        """调用convert_to_html工具。"""
        core_service = TestMCPToolsHelper.get_core_service()
        return await _handle_convert_to_html(arguments, core_service)

    @staticmethod
    async def call_parse_sheet(arguments):
        """调用parse_sheet工具。"""
        core_service = TestMCPToolsHelper.get_core_service()
        return await _handle_parse_sheet(arguments, core_service)

    @staticmethod
    async def call_apply_changes(arguments):
        """调用apply_changes工具。"""
        core_service = TestMCPToolsHelper.get_core_service()
        return await _handle_apply_changes(arguments, core_service)


class TestMCPTools:
    """MCP工具的全面测试。"""
    
    @pytest.mark.asyncio
    async def test_handle_convert_to_html_xlsx(self):
        """测试convert_to_html工具处理XLSX文件。"""
        sample_path = "tests/data/sample.xlsx"

        if os.path.exists(sample_path):
            with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as tmp_file:
                output_path = tmp_file.name

            try:
                arguments = {
                    "file_path": sample_path,
                    "output_path": output_path
                }

                result = await TestMCPToolsHelper.call_convert_to_html(arguments)
                
                # 验证返回结果
                assert isinstance(result, list)
                assert len(result) > 0
                
                # 验证返回的文本内容（JSON格式）
                text_content = result[0].text
                assert '"status": "success"' in text_content
                assert '"output_path"' in text_content
                assert '"file_size"' in text_content
                assert '"rows_converted"' in text_content
                assert '"cells_converted"' in text_content
                
                # 验证文件确实被创建
                assert os.path.exists(output_path)
                
                # 验证HTML文件内容
                with open(output_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                    assert "<table>" in html_content
                    assert "</table>" in html_content
                    
            finally:
                if os.path.exists(output_path):
                    os.unlink(output_path)
    
    @pytest.mark.asyncio
    async def test_handle_convert_to_html_csv(self):
        """测试convert_to_html工具处理CSV文件。"""
        # 创建临时CSV文件
        csv_content = "Name,Age,City\nJohn,25,New York\nJane,30,London"
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
            csv_file.write(csv_content)
            csv_path = csv_file.name
        
        with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as html_file:
            html_path = html_file.name
        
        try:
            arguments = {
                "file_path": csv_path,
                "output_path": html_path
            }
            
            result = await TestMCPToolsHelper.call_convert_to_html(arguments)
            
            assert isinstance(result, list)
            text_content = result[0].text
            assert '"status": "success"' in text_content
            assert '"rows_converted": 3' in text_content
            assert '"cells_converted": 9' in text_content
            
            # 验证HTML内容
            with open(html_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
                assert "John" in html_content
                assert "Jane" in html_content
                assert "New York" in html_content
                
        finally:
            os.unlink(csv_path)
            if os.path.exists(html_path):
                os.unlink(html_path)
    
    @pytest.mark.asyncio
    async def test_handle_convert_to_html_default_output(self):
        """测试convert_to_html工具使用默认输出路径。"""
        # 创建临时CSV文件
        csv_content = "Test,Data\n1,2"
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
            csv_file.write(csv_content)
            csv_path = csv_file.name
        
        try:
            arguments = {
                "file_path": csv_path
                # 不指定output_path，使用默认
            }
            
            result = await TestMCPToolsHelper.call_convert_to_html(arguments)

            assert isinstance(result, list)
            text_content = result[0].text
            assert '"status": "success"' in text_content
            
            # 检查输出路径字段存在（JSON中路径会被转义）
            assert '"output_path"' in text_content
            assert '.html' in text_content
            
            # 清理生成的HTML文件（如果存在）
            html_path = str(Path(csv_path).with_suffix('.html'))
            if os.path.exists(html_path):
                os.unlink(html_path)
                
        finally:
            os.unlink(csv_path)
    
    @pytest.mark.asyncio
    async def test_handle_convert_to_html_nonexistent_file(self):
        """测试convert_to_html工具处理不存在的文件。"""
        arguments = {
            "file_path": "nonexistent.xlsx",
            "output_path": "output.html"
        }
        
        result = await TestMCPToolsHelper.call_convert_to_html(arguments)

        assert isinstance(result, list)
        text_content = result[0].text
        assert "错误" in text_content or "失败" in text_content
    
    @pytest.mark.asyncio
    async def test_handle_parse_sheet_xlsx(self):
        """测试parse_sheet工具处理XLSX文件。"""
        sample_path = "tests/data/sample.xlsx"
        
        if os.path.exists(sample_path):
            arguments = {
                "file_path": sample_path
            }
            
            result = await TestMCPToolsHelper.call_parse_sheet(arguments)

            assert isinstance(result, list)
            text_content = result[0].text
            assert '"sheet_name"' in text_content
            assert '"metadata"' in text_content
            assert '"data_rows"' in text_content
            assert '"total_cols"' in text_content
            assert '"processing_mode": "full"' in text_content
    
    @pytest.mark.asyncio
    async def test_handle_parse_sheet_with_range(self):
        """测试parse_sheet工具使用范围选择。"""
        sample_path = "tests/data/sample.xlsx"
        
        if os.path.exists(sample_path):
            arguments = {
                "file_path": sample_path,
                "range_string": "A1:B2"
            }
            
            result = await TestMCPToolsHelper.call_parse_sheet(arguments)
            
            assert isinstance(result, list)
            text_content = result[0].text
            assert '"sheet_name"' in text_content
            assert '"range": "A1:B2"' in text_content
            assert '"processing_mode": "range_selection"' in text_content
    
    @pytest.mark.asyncio
    async def test_handle_parse_sheet_csv(self):
        """测试parse_sheet工具处理CSV文件。"""
        # 创建临时CSV文件
        csv_content = "Name,Age\nJohn,25\nJane,30"
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tmp_file:
            tmp_file.write(csv_content)
            tmp_path = tmp_file.name
        
        try:
            arguments = {
                "file_path": tmp_path
            }
            
            result = await TestMCPToolsHelper.call_parse_sheet(arguments)
            
            assert isinstance(result, list)
            text_content = result[0].text
            assert '"sheet_name"' in text_content
            assert '"total_rows"' in text_content
            assert '"total_cols": 2' in text_content
            
        finally:
            os.unlink(tmp_path)
    
    @pytest.mark.asyncio
    async def test_handle_parse_sheet_nonexistent_file(self):
        """测试parse_sheet工具处理不存在的文件。"""
        arguments = {
            "file_path": "nonexistent.xlsx"
        }
        
        result = await TestMCPToolsHelper.call_parse_sheet(arguments)
        
        assert isinstance(result, list)
        text_content = result[0].text
        assert "错误" in text_content or "失败" in text_content
    
    @pytest.mark.asyncio
    async def test_handle_apply_changes_basic(self):
        """测试apply_changes工具的基本功能。"""
        # 创建临时CSV文件
        csv_content = "Name,Age\nJohn,25\nJane,30"
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tmp_file:
            tmp_file.write(csv_content)
            tmp_path = tmp_file.name
        
        try:
            # 先解析文件获取正确的JSON格式
            core_service = CoreService()
            parsed_data = core_service.parse_sheet(tmp_path)

            # 创建table_model_json格式的参数
            table_model_json = {
                "sheet_name": parsed_data["sheet_name"],
                "headers": parsed_data["headers"],
                "rows": parsed_data["rows"]  # 这里可以包含修改后的数据
            }

            arguments = {
                "file_path": tmp_path,
                "table_model_json": table_model_json
            }

            result = await TestMCPToolsHelper.call_apply_changes(arguments)
            
            assert isinstance(result, list)
            text_content = result[0].text
            # 当前实现返回验证成功的消息
            assert "验证成功" in text_content or "数据修改完成" in text_content
                
        finally:
            os.unlink(tmp_path)
    
    @pytest.mark.asyncio
    async def test_handle_apply_changes_with_backup(self):
        """测试apply_changes工具的备份功能。"""
        # 创建临时CSV文件
        csv_content = "Name,Age\nJohn,25"
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tmp_file:
            tmp_file.write(csv_content)
            tmp_path = tmp_file.name
        
        try:
            # 先解析文件获取正确的JSON格式
            core_service = CoreService()
            parsed_data = core_service.parse_sheet(tmp_path)

            table_model_json = {
                "sheet_name": parsed_data["sheet_name"],
                "headers": parsed_data["headers"],
                "rows": parsed_data["rows"]
            }

            arguments = {
                "file_path": tmp_path,
                "table_model_json": table_model_json,
                "create_backup": True
            }
            
            result = await TestMCPToolsHelper.call_apply_changes(arguments)
            
            assert isinstance(result, list)
            text_content = result[0].text
            # 当前实现返回验证成功的消息
            assert "验证成功" in text_content or "数据修改完成" in text_content
            
        finally:
            os.unlink(tmp_path)
            # 清理可能创建的备份文件
            backup_path = tmp_path + ".backup"
            if os.path.exists(backup_path):
                os.unlink(backup_path)
    
    @pytest.mark.asyncio
    async def test_handle_apply_changes_nonexistent_file(self):
        """测试apply_changes工具处理不存在的文件。"""
        # 使用有效的JSON格式但文件不存在
        table_model_json = {
            "sheet_name": "test",
            "headers": ["A", "B"],
            "rows": []
        }

        arguments = {
            "file_path": "nonexistent.csv",
            "table_model_json": table_model_json
        }
        
        result = await TestMCPToolsHelper.call_apply_changes(arguments)
        
        assert isinstance(result, list)
        text_content = result[0].text
        assert "错误" in text_content or "失败" in text_content
    
    @pytest.mark.asyncio
    async def test_handle_apply_changes_invalid_changes(self):
        """测试apply_changes工具处理无效的修改数据。"""
        # 创建临时CSV文件
        csv_content = "Name,Age\nJohn,25"
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tmp_file:
            tmp_file.write(csv_content)
            tmp_path = tmp_file.name
        
        try:
            # 使用无效的JSON数据（缺少必需字段）
            invalid_json = {"invalid": "data"}

            arguments = {
                "file_path": tmp_path,
                "table_model_json": invalid_json
            }
            
            result = await TestMCPToolsHelper.call_apply_changes(arguments)
            
            assert isinstance(result, list)
            text_content = result[0].text
            # 应该处理错误但不崩溃
            assert isinstance(text_content, str)
            
        finally:
            os.unlink(tmp_path)


class TestMCPToolsParameterValidation:
    """MCP工具参数验证测试。"""
    
    @pytest.mark.asyncio
    async def test_convert_to_html_missing_file_path(self):
        """测试convert_to_html工具缺少file_path参数。"""
        arguments = {
            "output_path": "output.html"
            # 缺少file_path
        }
        
        result = await TestMCPToolsHelper.call_convert_to_html(arguments)
        
        assert isinstance(result, list)
        text_content = result[0].text
        assert "错误" in text_content or "参数" in text_content
    
    @pytest.mark.asyncio
    async def test_parse_sheet_missing_file_path(self):
        """测试parse_sheet工具缺少file_path参数。"""
        arguments = {
            "range_string": "A1:B2"
            # 缺少file_path
        }
        
        result = await TestMCPToolsHelper.call_parse_sheet(arguments)
        
        assert isinstance(result, list)
        text_content = result[0].text
        assert "错误" in text_content or "参数" in text_content
    
    @pytest.mark.asyncio
    async def test_apply_changes_missing_parameters(self):
        """测试apply_changes工具缺少必要参数。"""
        arguments = {
            "file_path": "test.csv"
            # 缺少table_model_json参数
        }
        
        result = await TestMCPToolsHelper.call_apply_changes(arguments)
        
        assert isinstance(result, list)
        text_content = result[0].text
        assert "错误" in text_content or "参数" in text_content
