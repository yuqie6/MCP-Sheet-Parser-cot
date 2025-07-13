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
                description="将表格文件（如 CSV, XLSX, XLS 等）转换为高保真 HTML，保留大部分样式。对于大型文件，会自动进行分页以优化性能和阅读体验。返回包含HTML文件路径和转换摘要的JSON对象。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "源表格文件的绝对路径，支持 .csv, .xlsx, .xls, .xlsb, .xlsm 格式。"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "输出HTML文件的路径。如果留空，将在源文件目录中生成一个同名的 .html 文件。"
                        },
                        "page_size": {
                            "type": "integer",
                            "description": "【可选】分页时每页显示的行数。默认为100行。用于控制大型文件转换后HTML的单页大小。"
                        },
                        "page_number": {
                            "type": "integer",
                            "description": "【可选】要查看的页码，从1开始。默认为1。用于浏览大型文件的特定页面。"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="parse_sheet",
                description="将表格文件的指定工作表和范围解析为结构化的 TableModel JSON 对象。这个JSON格式对大型语言模型（LLM）非常友好，方便进行数据分析和修改。对于大型工作表，会自动返回摘要信息以避免过多的输出。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "目标表格文件的绝对路径，支持 .csv, .xlsx, .xls, .xlsb, .xlsm 格式。"
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "【可选】要解析的工作表的名称。如果留空，将自动使用活动工作表或第一个工作表。"
                        },
                        "range_string": {
                            "type": "string",
                            "description": "【可选】要解析的单元格范围，格式如 'A1:D10'。如果留空，将解析整个工作表。建议在处理大文件时指定此参数以提高效率。"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="apply_changes",
                description="将通过 `parse_sheet` 获取并已修改的 TableModel JSON 数据写回到原始表格文件中。此工具能够完成数据的闭环操作，同时保留原始文件的格式和大部分样式。默认会自动创建文件备份。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "需要写回数据的目标文件的绝对路径。"
                        },
                        "table_model_json": {
                            "type": "object",
                            "description": "从 `parse_sheet` 工具获取并修改后的 TableModel JSON 数据。",
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
                            "description": "【可选】是否在写入前创建原始文件的备份。默认为 `true`，以防意外覆盖。"
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
    if not isinstance(file_path, str):
        return [TextContent(
            type="text",
            text="错误: 必须提供有效的 file_path 参数"
        )]
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
    if not isinstance(file_path, str):
        return [TextContent(
            type="text",
            text="错误: 必须提供有效的 file_path 参数"
        )]
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
    if not isinstance(file_path, str):
        return [TextContent(
            type="text",
            text="错误: 必须提供有效的 file_path 参数"
        )]
    
    table_model_json = arguments.get("table_model_json")
    if not isinstance(table_model_json, dict):
        return [TextContent(
            type="text",
            text="错误: 必须提供有效的 table_model_json 参数"
        )]
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
