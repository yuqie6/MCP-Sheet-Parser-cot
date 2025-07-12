"""
MCP 工具定义模块 - 三个核心工具

基于systemPatterns.md设计，实现完整的表格处理闭环：
1. convert_to_html - 完美HTML转换
2. parse_sheet - JSON数据解析（LLM友好格式）
3. apply_changes - 数据写回（完成编辑闭环）
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
                name="convert_to_html",
                description="将表格文件转换为完美的HTML文件，支持CSV、XLSX、XLS、XLSB、XLSM格式。支持大文件分页处理。返回HTML文件路径，不传输内容避免上下文爆炸。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "源表格文件的绝对路径。支持.csv、.xlsx、.xls、.xlsb、.xlsm格式。"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "输出HTML文件路径。如果未提供，将在源文件同目录生成同名.html文件。"
                        },
                        "page_size": {
                            "type": "integer",
                            "description": "分页大小（每页行数），用于大文件处理（可选）"
                        },
                        "page_number": {
                            "type": "integer",
                            "description": "页码（从1开始），用于大文件分页显示（可选，默认为1）"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="parse_sheet",
                description="将表格文件解析为LLM友好的TableModel JSON格式。支持范围选择和工作表选择，智能大小检测，大文件返回摘要。平均样式保真度97.64%。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "目标表格文件的绝对路径。支持CSV、XLSX、XLS、XLSB、XLSM格式。"
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "要解析的工作表名称。如果未提供，使用第一个工作表。"
                        },
                        "range_string": {
                            "type": "string",
                            "description": "要解析的单元格范围，如'A1:D10'。如果未提供，解析整个工作表。大文件时建议指定范围。"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="apply_changes",
                description="将修改后的TableModel JSON数据写回原始文件，完成编辑闭环。自动备份原文件，保持文件格式和样式。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "要写回的目标文件绝对路径。"
                        },
                        "table_model_json": {
                            "type": "object",
                            "description": "修改后的TableModel JSON数据，来自parse_sheet工具的输出。",
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
                        },
                        "create_backup": {
                            "type": "boolean",
                            "description": "是否创建备份文件。默认为true。"
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
            if name == "convert_to_html":
                return await _handle_convert_to_html(arguments, core_service)
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


async def _handle_convert_to_html(arguments: dict[str, Any], core_service: CoreService) -> list[TextContent]:
    """处理 convert_to_html 工具调用。"""
    file_path = arguments.get("file_path")
    output_path = arguments.get("output_path")
    page_size = arguments.get("page_size")
    page_number = arguments.get("page_number")

    if not file_path:
        return [TextContent(
            type="text",
            text="错误: 必须提供 file_path 参数"
        )]

    try:
        result = core_service.convert_to_html(file_path, output_path, page_size, page_number)
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

    if not file_path:
        return [TextContent(
            type="text",
            text="错误: 必须提供 file_path 参数"
        )]

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
    create_backup = arguments.get("create_backup", True)

    try:
        result = core_service.apply_changes(file_path, table_model_json, create_backup)
        return [TextContent(
            type="text",
            text=json.dumps(result, ensure_ascii=False, indent=2)
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"错误: {str(e)}"
        )]
