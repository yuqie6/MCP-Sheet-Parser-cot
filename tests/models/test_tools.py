
import pytest
import json
from unittest.mock import MagicMock, AsyncMock, patch

from src.models.tools import register_tools, _handle_convert_to_html, _handle_parse_sheet, _handle_apply_changes, _generate_next_steps_guidance
from src.core_service import CoreService
from mcp.server import Server
from mcp.types import TextContent

@pytest.fixture
def mock_server():
    """Fixture for a mocked Server instance."""
    server = MagicMock(spec=Server)
    list_tools_decorator = MagicMock()
    call_tool_decorator = MagicMock()
    server.list_tools.return_value = list_tools_decorator
    server.call_tool.return_value = call_tool_decorator
    return server

@pytest.fixture
def mock_core_service():
    """Fixture for a mocked CoreService instance."""
    return MagicMock(spec=CoreService)

def test_register_tools(mock_server):
    """Test that tools are registered correctly."""
    register_tools(mock_server)
    mock_server.list_tools.assert_called_once()
    mock_server.call_tool.assert_called_once()
    mock_server.list_tools.return_value.assert_called_once()
    mock_server.call_tool.return_value.assert_called_once()

@pytest.mark.asyncio
async def test_handle_list_tools(mock_server):
    """Test the handle_list_tools function."""
    register_tools(mock_server)
    list_tools_func = mock_server.list_tools.return_value.call_args[0][0]
    tools = await list_tools_func()
    assert len(tools) == 3
    assert tools[0].name == "convert_to_html"
    assert tools[1].name == "parse_sheet"
    assert tools[2].name == "apply_changes"

@pytest.mark.asyncio
async def test_handle_call_tool_dispatch(mock_server, mock_core_service):
    """Test that handle_call_tool dispatches to the correct handler."""
    register_tools(mock_server)
    call_tool_func = mock_server.call_tool.return_value.call_args[0][0]

    with patch('src.models.tools._handle_convert_to_html', new_callable=AsyncMock) as mock_convert:
        await call_tool_func("convert_to_html", {})
        mock_convert.assert_called_once()

    with patch('src.models.tools._handle_parse_sheet', new_callable=AsyncMock) as mock_parse:
        await call_tool_func("parse_sheet", {})
        mock_parse.assert_called_once()

    with patch('src.models.tools._handle_apply_changes', new_callable=AsyncMock) as mock_apply:
        await call_tool_func("apply_changes", {})
        mock_apply.assert_called_once()

@pytest.mark.asyncio
async def test_handle_call_tool_unknown_tool(mock_server):
    """Test handle_call_tool with an unknown tool name."""
    register_tools(mock_server)
    call_tool_func = mock_server.call_tool.return_value.call_args[0][0]
    result = await call_tool_func("unknown_tool", {})
    assert len(result) == 1
    assert isinstance(result[0], TextContent)
    assert "未知工具: unknown_tool" in result[0].text

@pytest.mark.asyncio
async def test_handle_call_tool_exception(mock_server):
    """Test handle_call_tool exception handling."""
    register_tools(mock_server)
    call_tool_func = mock_server.call_tool.return_value.call_args[0][0]
    with patch('src.models.tools._handle_convert_to_html', new_callable=AsyncMock) as mock_convert:
        mock_convert.side_effect = Exception("Test Error")
        result = await call_tool_func("convert_to_html", {})
        assert len(result) == 1
        assert "错误: Test Error" in result[0].text

# Tests for _handle_convert_to_html
@pytest.mark.asyncio
async def test_handle_convert_to_html_success(mock_core_service):
    """Test _handle_convert_to_html success case."""
    mock_core_service.convert_to_html.return_value = [{"file_path": "test.html", "file_size_kb": 10}]
    args = {"file_path": "test.xlsx"}
    result = await _handle_convert_to_html(args, mock_core_service)
    response = json.loads(result[0].text)
    assert response["success"] is True
    assert response["operation"] == "convert_to_html"
    assert len(response["results"]) == 1

@pytest.mark.asyncio
async def test_handle_convert_to_html_file_not_found(mock_core_service):
    """Test _handle_convert_to_html with FileNotFoundError."""
    mock_core_service.convert_to_html.side_effect = FileNotFoundError("File not found")
    args = {"file_path": "nonexistent.xlsx"}
    result = await _handle_convert_to_html(args, mock_core_service)
    response = json.loads(result[0].text)
    assert response["success"] is False
    assert response["error_type"] == "file_not_found"

# Tests for _handle_parse_sheet
@pytest.mark.asyncio
async def test_handle_parse_sheet_success(mock_core_service):
    """Test _handle_parse_sheet success case."""
    mock_core_service.parse_sheet_optimized.return_value = {"metadata": {}, "preview_rows": []}
    args = {"file_path": "test.xlsx"}
    result = await _handle_parse_sheet(args, mock_core_service)
    response = json.loads(result[0].text)
    assert response["success"] is True
    assert response["operation"] == "parse_sheet"

@pytest.mark.asyncio
async def test_handle_parse_sheet_value_error(mock_core_service):
    """Test _handle_parse_sheet with ValueError."""
    args = {"file_path": "test.xlsx", "preview_rows": -1}
    result = await _handle_parse_sheet(args, mock_core_service)
    response = json.loads(result[0].text)
    assert response["success"] is False
    assert response["error_type"] == "invalid_parameter"

# Tests for _handle_apply_changes
@pytest.mark.asyncio
async def test_handle_apply_changes_success(mock_core_service):
    """Test _handle_apply_changes success case."""
    mock_core_service.apply_changes.return_value = {"status": "success"}
    args = {"file_path": "test.xlsx", "table_model_json": {"sheet_name": "Sheet1", "headers": [], "rows": []}}
    result = await _handle_apply_changes(args, mock_core_service)
    response = json.loads(result[0].text)
    assert response["success"] is True
    assert response["operation"] == "apply_changes"

@pytest.mark.asyncio
async def test_handle_apply_changes_permission_error(mock_core_service):
    """Test _handle_apply_changes with PermissionError."""
    mock_core_service.apply_changes.side_effect = PermissionError("Permission denied")
    args = {"file_path": "test.xlsx", "table_model_json": {"sheet_name": "Sheet1", "headers": [], "rows": []}}
    result = await _handle_apply_changes(args, mock_core_service)
    response = json.loads(result[0].text)
    assert response["success"] is False
    assert response["error_type"] == "permission_error"

# Tests for _generate_next_steps_guidance
def test_generate_next_steps_guidance():
    """Test _generate_next_steps_guidance logic."""
    # Case 1: Full data not included
    result_meta = {"metadata": {"total_rows": 100, "preview_rows": 5, "has_styles": False, "total_cells": 500}}
    guidance = _generate_next_steps_guidance(result_meta, False, False)
    assert "设置include_full_data=true获取完整数据" in guidance[0]

    # Case 2: Styles not included
    result_meta = {"metadata": {"total_rows": 10, "preview_rows": 10, "has_styles": True, "total_cells": 50}}
    guidance = _generate_next_steps_guidance(result_meta, True, False)
    assert "设置include_styles=true获取样式数据" in guidance[0]

    # Case 3: Large file
    result_meta = {"metadata": {"total_rows": 10, "preview_rows": 10, "has_styles": False, "total_cells": 2000}}
    guidance = _generate_next_steps_guidance(result_meta, True, True)
    assert "建议使用range_string参数获取特定范围" in guidance[0]

    # Case 4: All data loaded
    result_meta = {"metadata": {"total_rows": 10, "preview_rows": 10, "has_styles": False, "total_cells": 50}}
    guidance = _generate_next_steps_guidance(result_meta, True, True)
    assert "数据已完整加载" in guidance[0]

class TestConvertToHtmlExceptionHandling:
    """测试convert_to_html的异常处理。"""

    @pytest.mark.asyncio
    async def test_handle_convert_to_html_permission_error(self, mock_core_service):
        """
        TDD测试：_handle_convert_to_html应该处理PermissionError

        这个测试覆盖第205-214行的PermissionError处理代码
        """
        mock_core_service.convert_to_html.side_effect = PermissionError("Permission denied")

        arguments = {
            "file_path": "test.xlsx",
            "output_path": "output.html"
        }

        result = await _handle_convert_to_html(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        assert isinstance(result[0], TextContent)
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "permission_error"
        assert "权限不足" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_convert_to_html_value_error(self, mock_core_service):
        """
        TDD测试：_handle_convert_to_html应该处理ValueError

        这个测试覆盖第215-224行的ValueError处理代码
        """
        mock_core_service.convert_to_html.side_effect = ValueError("Invalid parameter")

        arguments = {
            "file_path": "test.xlsx",
            "output_path": "output.html"
        }

        result = await _handle_convert_to_html(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        assert isinstance(result[0], TextContent)
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "invalid_parameter"
        assert "参数错误" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_convert_to_html_general_exception(self, mock_core_service):
        """
        TDD测试：_handle_convert_to_html应该处理一般异常

        这个测试覆盖第225行之后的一般异常处理代码
        """
        mock_core_service.convert_to_html.side_effect = RuntimeError("Unexpected error")

        arguments = {
            "file_path": "test.xlsx",
            "output_path": "output.html"
        }

        result = await _handle_convert_to_html(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        assert isinstance(result[0], TextContent)
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "conversion_error"

class TestParseSheetParameterValidation:
    """测试parse_sheet的参数验证。"""

    @pytest.mark.asyncio
    async def test_handle_parse_sheet_invalid_file_path(self, mock_core_service):
        """
        TDD测试：_handle_parse_sheet应该验证file_path参数

        这个测试覆盖第244行的file_path验证代码
        """
        arguments = {
            "file_path": ""  # 空字符串
        }

        result = await _handle_parse_sheet(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "invalid_parameter"
        assert "file_path必须是非空字符串" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_parse_sheet_invalid_sheet_name(self, mock_core_service):
        """
        TDD测试：_handle_parse_sheet应该验证sheet_name参数

        这个测试覆盖第248行的sheet_name验证代码
        """
        arguments = {
            "file_path": "test.xlsx",
            "sheet_name": 123  # 非字符串
        }

        result = await _handle_parse_sheet(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "invalid_parameter"
        assert "sheet_name必须是字符串" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_parse_sheet_invalid_range_string(self, mock_core_service):
        """
        TDD测试：_handle_parse_sheet应该验证range_string参数

        这个测试覆盖第252行的range_string验证代码
        """
        arguments = {
            "file_path": "test.xlsx",
            "range_string": 123  # 非字符串
        }

        result = await _handle_parse_sheet(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "invalid_parameter"
        assert "range_string必须是字符串" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_parse_sheet_invalid_include_full_data(self, mock_core_service):
        """
        TDD测试：_handle_parse_sheet应该验证include_full_data参数

        这个测试覆盖第256行的include_full_data验证代码
        """
        arguments = {
            "file_path": "test.xlsx",
            "include_full_data": "true"  # 非布尔值
        }

        result = await _handle_parse_sheet(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "invalid_parameter"
        assert "include_full_data必须是布尔值" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_parse_sheet_invalid_include_styles(self, mock_core_service):
        """
        TDD测试：_handle_parse_sheet应该验证include_styles参数

        这个测试覆盖第260行的include_styles验证代码
        """
        arguments = {
            "file_path": "test.xlsx",
            "include_styles": "false"  # 非布尔值
        }

        result = await _handle_parse_sheet(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "invalid_parameter"
        assert "include_styles必须是布尔值" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_parse_sheet_invalid_max_rows(self, mock_core_service):
        """
        TDD测试：_handle_parse_sheet应该验证max_rows参数

        这个测试覆盖第268行的max_rows验证代码
        """
        arguments = {
            "file_path": "test.xlsx",
            "max_rows": -1  # 负数
        }

        result = await _handle_parse_sheet(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "invalid_parameter"
        assert "max_rows必须是正整数或None" in response["error_message"]

class TestParseSheetExceptionHandling:
    """测试parse_sheet的异常处理。"""

    @pytest.mark.asyncio
    async def test_handle_parse_sheet_file_not_found_error(self, mock_core_service):
        """
        TDD测试：_handle_parse_sheet应该处理FileNotFoundError

        这个测试覆盖第304行的FileNotFoundError处理代码
        """
        mock_core_service.parse_sheet_optimized.side_effect = FileNotFoundError("File not found")

        arguments = {
            "file_path": "nonexistent.xlsx"
        }

        result = await _handle_parse_sheet(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "file_not_found"
        assert "文件未找到" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_parse_sheet_general_exception(self, mock_core_service):
        """
        TDD测试：_handle_parse_sheet应该处理一般异常

        这个测试覆盖第323-324行的一般异常处理代码
        """
        mock_core_service.parse_sheet_optimized.side_effect = RuntimeError("Parsing failed")

        arguments = {
            "file_path": "test.xlsx"
        }

        result = await _handle_parse_sheet(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "parsing_error"
        assert "解析失败" in response["error_message"]

class TestApplyChangesExceptionHandling:
    """测试apply_changes的异常处理。"""

    @pytest.mark.asyncio
    async def test_handle_apply_changes_file_not_found_error(self, mock_core_service):
        """
        TDD测试：_handle_apply_changes应该处理FileNotFoundError

        这个测试覆盖第380行的FileNotFoundError处理代码
        """
        mock_core_service.apply_changes.side_effect = FileNotFoundError("File not found")

        arguments = {
            "file_path": "nonexistent.xlsx",
            "table_model_json": '{"sheet_name": "Sheet1", "headers": [], "rows": []}'
        }

        result = await _handle_apply_changes(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "file_not_found"
        assert "文件未找到" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_apply_changes_value_error(self, mock_core_service):
        """
        TDD测试：_handle_apply_changes应该处理ValueError

        这个测试覆盖第399-408行的ValueError处理代码
        """
        mock_core_service.apply_changes.side_effect = ValueError("Invalid data format")

        arguments = {
            "file_path": "test.xlsx",
            "table_model_json": '{"invalid": "data"}'
        }

        result = await _handle_apply_changes(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "invalid_data"
        assert "数据格式错误" in response["error_message"]

    @pytest.mark.asyncio
    async def test_handle_apply_changes_general_exception(self, mock_core_service):
        """
        TDD测试：_handle_apply_changes应该处理一般异常

        这个测试覆盖第409-410行之后的一般异常处理代码
        """
        mock_core_service.apply_changes.side_effect = RuntimeError("Apply failed")

        arguments = {
            "file_path": "test.xlsx",
            "table_model_json": '{"sheet_name": "Sheet1", "headers": [], "rows": []}'
        }

        result = await _handle_apply_changes(arguments, mock_core_service)

        # 验证返回的错误响应
        assert len(result) == 1
        response = json.loads(result[0].text)
        assert response["success"] is False
        assert response["error_type"] == "write_error"
