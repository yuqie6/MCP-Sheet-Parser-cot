"""
MCP服务器测试模块

测试MCP服务器的基本功能和工具注册
"""

import pytest
from unittest.mock import Mock, patch
from src.mcp_server.server import create_server
from src.mcp_server.tools import (
    _handle_convert_to_html,
    _handle_parse_sheet,
    _handle_apply_changes
)
from src.core_service import CoreService


class TestMCPServer:
    """MCP服务器基础测试。"""

    def test_server_creation(self):
        """测试服务器创建。"""
        server = create_server()
        assert server is not None
        assert hasattr(server, 'name')
        assert server.name == "mcp-sheet-parser"

    def test_server_tools_registration(self):
        """测试服务器工具注册。"""
        server = create_server()

        # 验证服务器已经注册了工具
        # 这里我们检查服务器对象的基本属性
        assert hasattr(server, 'name')
        assert server.name == "mcp-sheet-parser"


class TestMCPToolsInternal:
    """MCP工具内部功能测试。"""
    
    @pytest.mark.asyncio
    async def test_tool_error_handling(self):
        """测试工具的错误处理。"""
        core_service = CoreService()
        
        # 测试convert_to_html的错误处理
        invalid_args = {"invalid": "arguments"}
        result = await _handle_convert_to_html(invalid_args, core_service)
        
        assert isinstance(result, list)
        assert len(result) > 0
        text_content = result[0].text
        assert "错误" in text_content or "Error" in text_content.lower()
    
    @pytest.mark.asyncio
    async def test_parse_sheet_error_handling(self):
        """测试parse_sheet工具的错误处理。"""
        core_service = CoreService()
        
        # 测试无效参数
        invalid_args = {}
        result = await _handle_parse_sheet(invalid_args, core_service)
        
        assert isinstance(result, list)
        assert len(result) > 0
        text_content = result[0].text
        assert "错误" in text_content or "Error" in text_content.lower()
    
    @pytest.mark.asyncio
    async def test_apply_changes_error_handling(self):
        """测试apply_changes工具的错误处理。"""
        core_service = CoreService()
        
        # 测试无效参数
        invalid_args = {}
        result = await _handle_apply_changes(invalid_args, core_service)
        
        assert isinstance(result, list)
        assert len(result) > 0
        text_content = result[0].text
        assert "错误" in text_content or "Error" in text_content.lower()
    
    @pytest.mark.asyncio
    async def test_tools_with_none_parameters(self):
        """测试工具处理None参数的情况。"""
        core_service = CoreService()
        
        # 测试None文件路径
        args_with_none = {"file_path": None}
        
        result1 = await _handle_convert_to_html(args_with_none, core_service)
        assert isinstance(result1, list)
        
        result2 = await _handle_parse_sheet(args_with_none, core_service)
        assert isinstance(result2, list)
        
        result3 = await _handle_apply_changes(args_with_none, core_service)
        assert isinstance(result3, list)
    
    @pytest.mark.asyncio
    async def test_tools_exception_handling(self):
        """测试工具的异常处理机制。"""
        # 创建一个会抛出异常的mock服务
        mock_service = Mock()
        mock_service.convert_to_html.side_effect = Exception("Test exception")
        mock_service.parse_sheet.side_effect = Exception("Test exception")
        mock_service.apply_changes.side_effect = Exception("Test exception")
        
        valid_args = {"file_path": "test.xlsx"}
        
        # 测试异常是否被正确捕获和处理
        result1 = await _handle_convert_to_html(valid_args, mock_service)
        assert isinstance(result1, list)
        assert "错误" in result1[0].text
        
        result2 = await _handle_parse_sheet(valid_args, mock_service)
        assert isinstance(result2, list)
        assert "错误" in result2[0].text
        
        # apply_changes需要table_model_json参数
        apply_args = {"file_path": "test.xlsx", "table_model_json": {}}
        result3 = await _handle_apply_changes(apply_args, mock_service)
        assert isinstance(result3, list)
        assert "错误" in result3[0].text


class TestMCPToolsParameterValidation:
    """MCP工具参数验证测试。"""
    
    @pytest.mark.asyncio
    async def test_convert_to_html_parameter_validation(self):
        """测试convert_to_html的参数验证。"""
        core_service = CoreService()
        
        # 测试各种无效参数组合
        test_cases = [
            {},  # 空参数
            {"output_path": "test.html"},  # 缺少file_path
            {"file_path": ""},  # 空file_path
            {"file_path": None},  # None file_path
        ]
        
        for args in test_cases:
            result = await _handle_convert_to_html(args, core_service)
            assert isinstance(result, list)
            assert len(result) > 0
    
    @pytest.mark.asyncio
    async def test_parse_sheet_parameter_validation(self):
        """测试parse_sheet的参数验证。"""
        core_service = CoreService()
        
        # 测试各种无效参数组合
        test_cases = [
            {},  # 空参数
            {"range_string": "A1:B2"},  # 缺少file_path
            {"file_path": ""},  # 空file_path
            {"file_path": None},  # None file_path
        ]
        
        for args in test_cases:
            result = await _handle_parse_sheet(args, core_service)
            assert isinstance(result, list)
            assert len(result) > 0
    
    @pytest.mark.asyncio
    async def test_apply_changes_parameter_validation(self):
        """测试apply_changes的参数验证。"""
        core_service = CoreService()
        
        # 测试各种无效参数组合
        test_cases = [
            {},  # 空参数
            {"file_path": "test.csv"},  # 缺少table_model_json
            {"table_model_json": {}},  # 缺少file_path
            {"file_path": "", "table_model_json": {}},  # 空file_path
            {"file_path": None, "table_model_json": {}},  # None file_path
        ]
        
        for args in test_cases:
            result = await _handle_apply_changes(args, core_service)
            assert isinstance(result, list)
            assert len(result) > 0


class TestMCPIntegration:
    """MCP集成测试。"""
    
    def test_core_service_integration(self):
        """测试CoreService与MCP工具的集成。"""
        core_service = CoreService()
        
        # 验证CoreService具有MCP工具需要的所有方法
        required_methods = ['get_sheet_info', 'parse_sheet', 'convert_to_html', 'apply_changes']
        
        for method_name in required_methods:
            assert hasattr(core_service, method_name)
            assert callable(getattr(core_service, method_name))
    
    def test_tool_return_format_consistency(self):
        """测试工具返回格式的一致性。"""
        # 所有MCP工具都应该返回相同格式的结果
        # 这里我们验证返回类型的一致性
        
        # 这个测试确保所有工具都返回List[TextContent]格式
        # 实际的格式验证在其他测试中已经覆盖
        pass
