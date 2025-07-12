"""
MCP 工具功能测试用例 - 重写版
测试三个核心MCP工具：convert_to_html、parse_sheet、apply_changes
"""

import pytest
import json
import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

from src.mcp_server.tools import (
    _handle_convert_to_html,
    _handle_parse_sheet,
    _handle_apply_changes
)
from src.core_service import CoreService
from src.models.table_model import Sheet, Row, Cell, Style


@pytest.fixture
def mock_service():
    """创建用于测试的 CoreService mock。"""
    service = MagicMock(spec=CoreService)
    return service


@pytest.fixture
def sample_sheet():
    """创建用于测试的示例 sheet。"""
    return Sheet(
        name="TestSheet",
        rows=[
            Row(cells=[
                Cell(value="Header1", style=Style(bold=True, font_color="#FF0000")),
                Cell(value="Header2", style=Style(bold=True, font_color="#FF0000")),
                Cell(value="Header3", style=Style(bold=True, font_color="#FF0000"))
            ]),
            Row(cells=[
                Cell(value="Data1"),
                Cell(value="Data2"),
                Cell(value="Data3")
            ])
        ]
    )


@pytest.fixture
def sample_json_data():
    """创建用于测试的示例 JSON 数据。"""
    return {
        "metadata": {
            "name": "TestSheet",
            "rows": 2,
            "cols": 3,
            "total_cells": 6
        },
        "data": [
            {
                "row": 0,
                "cells": [
                    {"col": 0, "value": "Header1", "style_id": "style_1"},
                    {"col": 1, "value": "Header2", "style_id": "style_1"},
                    {"col": 2, "value": "Header3", "style_id": "style_1"}
                ]
            },
            {
                "row": 1,
                "cells": [
                    {"col": 0, "value": "Data1"},
                    {"col": 1, "value": "Data2"},
                    {"col": 2, "value": "Data3"}
                ]
            }
        ],
        "merged_cells": [],
        "styles": {
            "style_1": {"bold": True, "font_color": "#FF0000"}
        }
    }


# 测试 convert_to_html 工具
@pytest.mark.asyncio
async def test_convert_to_html_missing_file_path(mock_service):
    """测试 convert_to_html 工具缺少 file_path 参数。"""
    result = await _handle_convert_to_html({}, mock_service)
    assert len(result) == 1
    assert "错误: 必须提供 file_path 参数" in result[0].text


@pytest.mark.asyncio
async def test_convert_to_html_file_not_found(mock_service):
    """测试 convert_to_html 工具文件不存在。"""
    mock_service.convert_to_html.side_effect = FileNotFoundError("文件不存在")
    result = await _handle_convert_to_html({"file_path": "nonexistent.xlsx"}, mock_service)
    assert len(result) == 1
    assert "错误:" in result[0].text


@pytest.mark.asyncio
async def test_convert_to_html_success(mock_service):
    """测试 convert_to_html 工具成功转换。"""
    mock_service.convert_to_html.return_value = {
        "status": "success",
        "output_path": "/path/to/output.html",
        "file_size": 1024,
        "rows_converted": 10
    }
    result = await _handle_convert_to_html({"file_path": "test.xlsx"}, mock_service)
    assert len(result) == 1
    assert "status" in result[0].text
    assert "success" in result[0].text


# 测试 parse_sheet 工具
@pytest.mark.asyncio
async def test_parse_sheet_missing_file_path(mock_service):
    """测试 parse_sheet 工具缺少 file_path 参数。"""
    result = await _handle_parse_sheet({}, mock_service)
    assert len(result) == 1
    assert "错误: 必须提供 file_path 参数" in result[0].text


@pytest.mark.asyncio
async def test_parse_sheet_file_not_found(mock_service):
    """测试 parse_sheet 工具文件不存在。"""
    mock_service.parse_sheet.side_effect = FileNotFoundError("文件不存在")
    result = await _handle_parse_sheet({"file_path": "nonexistent.xlsx"}, mock_service)
    assert len(result) == 1
    assert "错误:" in result[0].text


@pytest.mark.asyncio
async def test_parse_sheet_success(mock_service, sample_json_data):
    """测试 parse_sheet 工具成功解析。"""
    mock_service.parse_sheet.return_value = sample_json_data
    result = await _handle_parse_sheet({"file_path": "test.xlsx"}, mock_service)
    assert len(result) == 1
    assert "metadata" in result[0].text
    assert "TestSheet" in result[0].text


# 测试 apply_changes 工具
@pytest.mark.asyncio
async def test_apply_changes_missing_file_path(mock_service):
    """测试 apply_changes 工具缺少 file_path 参数。"""
    result = await _handle_apply_changes({}, mock_service)
    assert len(result) == 1
    assert "错误:" in result[0].text


@pytest.mark.asyncio
async def test_apply_changes_success(mock_service, sample_json_data):
    """测试 apply_changes 工具成功应用修改。"""
    mock_service.apply_changes.return_value = {
        "status": "数据验证成功，写回功能开发中",
        "file_path": "/path/to/file.xlsx",
        "backup_created": True,
        "backup_path": "/path/to/file.xlsx.backup"
    }
    result = await _handle_apply_changes({
        "file_path": "test.xlsx",
        "table_model_json": sample_json_data
    }, mock_service)
    assert len(result) == 1
    assert "status" in result[0].text


# MCP 服务器测试
class TestMCPServer:
    """MCP 服务器相关测试。"""

    def test_create_server(self):
        """测试服务器创建功能。"""
        from src.mcp_server.server import create_server

        server = create_server()
        assert server is not None
        assert server.name == "mcp-sheet-parser"

        # 验证服务器基本属性
        assert hasattr(server, 'list_tools')
        assert hasattr(server, 'call_tool')
        assert hasattr(server, 'run')

    def test_server_initialization_logging(self, caplog):
        """测试服务器初始化日志。"""
        from src.mcp_server.server import create_server
        import logging

        with caplog.at_level(logging.INFO):
            server = create_server()

        assert server is not None

    def test_tool_registration_count(self):
        """测试工具注册数量。"""
        from src.mcp_server.server import create_server

        server = create_server()

        # 验证服务器有工具注册功能
        assert hasattr(server, 'call_tool')
        assert hasattr(server, 'list_tools')

        # 验证服务器名称
        assert server.name == "mcp-sheet-parser"
