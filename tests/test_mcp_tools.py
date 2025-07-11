"""
MCP 工具功能测试用例。
"""

import pytest
import json
import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

from src.mcp_server.tools import (
    _handle_parse_sheet_to_json,
    _handle_convert_json_to_html,
    _handle_convert_file_to_html,
    _handle_convert_file_to_html_file,
    _handle_get_table_summary,
    _handle_get_sheet_metadata,
    _json_to_sheet
)
from src.services.sheet_service import SheetService
from src.models.table_model import Sheet, Row, Cell, Style


@pytest.fixture
def mock_service():
    """创建用于测试的 SheetService mock。"""
    service = MagicMock(spec=SheetService)
    return service


@pytest.fixture
def sample_sheet():
    """创建用于测试的示例 sheet。"""
    return Sheet(
        name="TestSheet",
        rows=[
            Row(cells=[
                Cell(value="Header1", style=Style(bold=True, font_color="#FF0000")),
                Cell(value="Header2", style=Style(bold=True, font_color="#FF0000"))
            ]),
            Row(cells=[
                Cell(value="Data1"),
                Cell(value="Data2")
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
            "cols": 2,
            "has_merged_cells": False,
            "merged_cells_count": 0
        },
        "data": [
            {
                "row": 0,
                "cells": [
                    {"col": 0, "value": "Header1", "style_id": "style_1"},
                    {"col": 1, "value": "Header2", "style_id": "style_1"}
                ]
            },
            {
                "row": 1,
                "cells": [
                    {"col": 0, "value": "Data1", "style_id": None},
                    {"col": 1, "value": "Data2", "style_id": None}
                ]
            }
        ],
        "merged_cells": [],
        "styles": {
            "style_1": {
                "bold": True,
                "font_color": "#FF0000"
            }
        }
    }


@pytest.mark.asyncio
async def test_parse_sheet_to_json(mock_service):
    """测试 parse_sheet_to_json 工具。"""
    # 测试缺少 file_path
    with pytest.raises(ValueError, match="必须提供 file_path"):
        await _handle_parse_sheet_to_json({}, mock_service)
    
    # 测试不存在的文件
    with pytest.raises(FileNotFoundError):
        await _handle_parse_sheet_to_json({"file_path": "nonexistent.xlsx"}, mock_service)
    
    # 测试成功解析
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        result = await _handle_parse_sheet_to_json({"file_path": test_file}, mock_service)
        assert len(result) == 1
        assert "JSON 转换成功" in result[0].text
        assert "JSON 数据:" in result[0].text


@pytest.mark.asyncio
async def test_convert_json_to_html(mock_service, sample_json_data):
    """测试 convert_json_to_html 工具。"""
    # 测试缺少参数
    with pytest.raises(ValueError, match="必须提供 json_data"):
        await _handle_convert_json_to_html({}, mock_service)

    with pytest.raises(ValueError, match="必须提供 output_path"):
        await _handle_convert_json_to_html({"json_data": sample_json_data}, mock_service)
    
    # 测试成功转换
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
        output_path = tmp.name
    
    try:
        result = await _handle_convert_json_to_html({
            "json_data": sample_json_data,
            "output_path": output_path
        }, mock_service)
        
        assert len(result) == 1
        assert "HTML 文件生成成功" in result[0].text
        assert output_path in result[0].text
        assert Path(output_path).exists()
        
        # 验证 HTML 内容
        html_content = Path(output_path).read_text(encoding='utf-8')
        assert "TestSheet" in html_content
        assert "Header1" in html_content
        assert "Data1" in html_content
    finally:
        Path(output_path).unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_convert_file_to_html(mock_service):
    """测试 convert_file_to_html 工具。"""
    # 测试缺少 file_path
    with pytest.raises(ValueError, match="必须提供 file_path"):
        await _handle_convert_file_to_html({}, mock_service)
    
    # 测试不存在的文件
    with pytest.raises(FileNotFoundError):
        await _handle_convert_file_to_html({"file_path": "nonexistent.xlsx"}, mock_service)
    
    # 测试成功转换
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        result = await _handle_convert_file_to_html({"file_path": test_file}, mock_service)
        assert len(result) == 1
        assert "HTML 转换完成" in result[0].text


@pytest.mark.asyncio
async def test_convert_file_to_html_file(mock_service):
    """测试 convert_file_to_html_file 工具。"""
    # 测试缺少参数
    with pytest.raises(ValueError, match="必须提供 file_path"):
        await _handle_convert_file_to_html_file({}, mock_service)

    with pytest.raises(ValueError, match="必须提供 output_path"):
        await _handle_convert_file_to_html_file({"file_path": "test.xlsx"}, mock_service)
    
    # 测试成功转换
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
            output_path = tmp.name
        
        try:
            result = await _handle_convert_file_to_html_file({
                "file_path": test_file,
                "output_path": output_path
            }, mock_service)
            
            assert len(result) == 1
            assert "HTML 文件转换成功" in result[0].text
            assert Path(output_path).exists()
        finally:
            Path(output_path).unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_get_table_summary(mock_service):
    """测试 get_table_summary 工具。"""
    # 测试缺少 file_path
    with pytest.raises(ValueError, match="必须提供 file_path"):
        await _handle_get_table_summary({}, mock_service)
    
    # 测试不存在的文件
    with pytest.raises(FileNotFoundError):
        await _handle_get_table_summary({"file_path": "nonexistent.xlsx"}, mock_service)
    
    # 测试成功摘要
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        result = await _handle_get_table_summary({"file_path": test_file}, mock_service)
        assert len(result) == 1
        assert "表格摘要:" in result[0].text
        assert "基本统计:" in result[0].text
        assert "前5行样例:" in result[0].text
        assert "处理建议:" in result[0].text


@pytest.mark.asyncio
async def test_get_sheet_metadata(mock_service):
    """测试 get_sheet_metadata 工具。"""
    # 测试缺少 file_path
    with pytest.raises(ValueError, match="必须提供 file_path"):
        await _handle_get_sheet_metadata({}, mock_service)
    
    # 测试不存在的文件
    with pytest.raises(FileNotFoundError):
        await _handle_get_sheet_metadata({"file_path": "nonexistent.xlsx"}, mock_service)
    
    # 测试成功获取元数据
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        result = await _handle_get_sheet_metadata({"file_path": test_file}, mock_service)
        assert len(result) == 1
        assert "表格元数据:" in result[0].text
        assert "文件信息:" in result[0].text
        assert "结构:" in result[0].text
        assert "样式:" in result[0].text
        assert "输出估算:" in result[0].text


def test_json_to_sheet_conversion(sample_json_data):
    """测试 JSON 到 Sheet 的转换。"""
    sheet = _json_to_sheet(sample_json_data)
    
    # 测试基本属性
    assert sheet.name == "TestSheet"
    assert len(sheet.rows) == 2
    assert len(sheet.rows[0].cells) == 2
    
    # 测试单元格值
    assert sheet.rows[0].cells[0].value == "Header1"
    assert sheet.rows[0].cells[1].value == "Header2"
    assert sheet.rows[1].cells[0].value == "Data1"
    assert sheet.rows[1].cells[1].value == "Data2"
    
    # 测试样式
    assert sheet.rows[0].cells[0].style is not None
    assert sheet.rows[0].cells[0].style.bold == True
    assert sheet.rows[0].cells[0].style.font_color == "#FF0000"
    assert sheet.rows[1].cells[0].style is None  # No style for data cells


def test_json_to_sheet_empty_data():
    """测试空数据的 JSON 到 Sheet 转换。"""
    empty_json = {
        "metadata": {"name": "Empty", "rows": 0, "cols": 0},
        "data": [],
        "merged_cells": [],
        "styles": {}
    }
    
    sheet = _json_to_sheet(empty_json)
    assert sheet.name == "Empty"
    assert len(sheet.rows) == 0
    assert len(sheet.merged_cells) == 0


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

        assert "MCP 表格解析服务器初始化完成" in caplog.text
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

    @pytest.mark.asyncio
    async def test_parse_sheet_missing_file_path(self):
        """测试解析表格缺少文件路径参数。"""
        from src.mcp_server.tools import _handle_parse_sheet_to_json

        mock_service = MagicMock()

        # 测试缺少file_path参数，应该抛出ValueError
        with pytest.raises(ValueError, match="必须提供 file_path"):
            await _handle_parse_sheet_to_json({}, mock_service)

    @pytest.mark.asyncio
    async def test_parse_sheet_file_not_found(self):
        """测试解析不存在的文件。"""
        from src.mcp_server.tools import _handle_parse_sheet_to_json

        mock_service = MagicMock()

        # 测试文件不存在，应该抛出FileNotFoundError
        with pytest.raises(FileNotFoundError, match="文件未找到"):
            await _handle_parse_sheet_to_json(
                {"file_path": "nonexistent.xlsx"},
                mock_service
            )

    @pytest.mark.asyncio
    async def test_convert_json_to_html_missing_output_path(self):
        """测试JSON转HTML缺少输出路径。"""
        from src.mcp_server.tools import _handle_convert_json_to_html

        mock_service = MagicMock()

        # 测试缺少output_path参数，应该抛出ValueError
        with pytest.raises(ValueError, match="必须提供 output_path"):
            await _handle_convert_json_to_html(
                {"json_data": "some json"},
                mock_service
            )
