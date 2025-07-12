"""
MCP 工具定义模块 - 简洁版

定义3个核心工具，实现完整的表格处理闭环：
1. get_sheet_info - 获取表格元数据
2. parse_sheet - 解析表格数据为标准JSON
3. apply_changes - 将修改应用回原始文件
"""

import logging
import json
from typing import Any
from pathlib import Path

from mcp.server import Server
from mcp.types import Tool, TextContent

from ..core_service import CoreService

logger = logging.getLogger(__name__)


def register_tools(server: Server) -> None:
    """向服务器注册3个核心MCP工具。"""

    # 初始化核心服务
    core_service = CoreService()
    
    @server.list_tools()
    async def handle_list_tools() -> list[Tool]:
        return [
            Tool(
                name="get_sheet_info",
                description="获取电子表格文件的元数据，如工作表名称、维度和总单元格数，而不加载完整数据。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "目标电子表格文件的绝对路径。"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="parse_sheet",
                description="将电子表格文件按指定范围解析为标准化的 TableModel JSON 格式。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "目标电子表格文件的绝对路径。"
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "要解析的目标工作表的名称。如果未提供，则默认为第一个工作表。"
                        },
                        "range_string": {
                            "type": "string",
                            "description": "要解析的单元格范围，格式如 'A1:D10'。如果未提供，则解析整个工作表。"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="apply_changes",
                description="将一个 TableModel JSON 对象的更改应用回原始的电子表格文件。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "要应用更改的目标电子表格文件的绝对路径。"
                        },
                        "table_model_json": {
                            "type": "object",
                            "description": "包含已修改数据的 TableModel JSON 对象。",
                            "properties": {
                                "sheet_name": {"type": "string"},
                                "headers": {
                                    "type": "array",
                                    "items": {"type": "string"}
                                },
                                "rows": {
                                    "type": "array",
                                    "items": {
                                        "type": "array",
                                        "items": {}
                                    }
                                }
                            },
                            "required": ["sheet_name", "headers", "rows"]
                        }
                    },
                    "required": ["file_path", "table_model_json"]
                }
            )
        ]

    @server.call_tool()
    async def handle_call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
        """处理工具调用请求。"""
        try:
            if name == "get_sheet_info":
                return await _handle_get_sheet_info(arguments, core_service)
            elif name == "parse_sheet":
                return await _handle_parse_sheet(arguments, core_service)
            elif name == "apply_changes":
                return await _handle_apply_changes(arguments, core_service)
            else:
                return [TextContent(
                    type="text",
                    text=f"未知工具: {name}"
                )]
        except Exception as e:
            logger.error(f"工具调用失败 {name}: {e}")
            return [TextContent(
                type="text",
                text=f"错误: {str(e)}"
            )]


async def _handle_get_sheet_info(arguments: dict[str, Any], core_service: CoreService) -> list[TextContent]:
    """处理 get_sheet_info 工具调用。"""
    file_path = arguments.get("file_path")

    try:
        result = core_service.get_sheet_info(file_path)
        return [TextContent(
            type="text",
            text=json.dumps(result, ensure_ascii=False, indent=2)
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"错误: {str(e)}"
        )]


async def _handle_parse_sheet(arguments: dict[str, Any], core_service: CoreService) -> list[TextContent]:
    """处理 parse_sheet 工具调用。"""
    file_path = arguments.get("file_path")
    sheet_name = arguments.get("sheet_name")
    range_string = arguments.get("range_string")

    try:
        result = core_service.parse_sheet(file_path, sheet_name, range_string)
        return [TextContent(
            type="text",
            text=json.dumps(result, ensure_ascii=False, indent=2)
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"错误: {str(e)}"
        )]


async def _handle_apply_changes(arguments: dict[str, Any], core_service: CoreService) -> list[TextContent]:
    """处理 apply_changes 工具调用。"""
    file_path = arguments.get("file_path")
    table_model_json = arguments.get("table_model_json")

    try:
        result = core_service.apply_changes(file_path, table_model_json)
        return [TextContent(
            type="text",
            text=json.dumps(result, ensure_ascii=False, indent=2)
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"错误: {str(e)}"
        )]
