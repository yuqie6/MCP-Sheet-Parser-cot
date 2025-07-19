
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
