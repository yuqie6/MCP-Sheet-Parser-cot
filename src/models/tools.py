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
                description="将Excel/CSV文件转换为可在浏览器中查看的HTML文件。保留原始样式、颜色、字体等格式。支持多工作表文件，可选择特定工作表或转换全部。大文件可使用分页功能。返回结构化JSON，包含成功状态、生成的文件信息和转换摘要。",
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
                        "sheet_name": {
                            "type": "string",
                            "description": "【可选】要转换的单个工作表的名称。如果留空，将转换文件中的所有工作表。"
                        },
                        "page_size": {
                            "type": "integer",
                            "description": "【可选】分页时每页显示的行数。默认为100行。用于控制大型文件转换后HTML的单页大小。"
                        },
                        "page_number": {
                            "type": "integer",
                            "description": "【可选】要查看的页码，从1开始。默认为1。用于浏览大型文件的特定页面。"
                        },
                        "header_rows": {
                            "type": "integer",
                            "description": "【可选】将文件顶部的指定行数视为表头。默认为 1。"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="parse_sheet",
                description="解析Excel/CSV文件为结构化JSON数据。默认返回文件概览信息（行数、列数、数据类型、前几行预览），避免上下文过载。LLM可通过参数控制是否获取完整数据、样式信息等。适合数据分析和处理，修改后可用apply_changes写回。",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "目标表格文件的绝对路径，支持 .csv, .xlsx, .xls, .xlsb, .xlsm 格式。"
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "【可选】要解析的工作表名称。如果留空，使用第一个工作表。"
                        },
                        "range_string": {
                            "type": "string",
                            "description": "【可选】单元格范围，如'A1:D10'。指定范围时会返回该范围的完整数据。"
                        },
                        "include_full_data": {
                            "type": "boolean",
                            "description": "【可选，默认false】是否返回完整数据。false时只返回概览和预览，true时返回所有行数据。大文件建议先查看概览。"
                        },
                        "include_styles": {
                            "type": "boolean",
                            "description": "【可选，默认false】是否包含样式信息（字体、颜色、边框等）。样式信息会显著增加数据量。"
                        },
                        "preview_rows": {
                            "type": "integer",
                            "description": "【可选，默认5】预览行数。当include_full_data为false时，返回前N行作为数据预览。"
                        },
                        "max_rows": {
                            "type": "integer",
                            "description": "【可选】最大返回行数。用于限制大文件的数据量，超出部分会被截断并提示。"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="apply_changes",
                description="将修改后的数据写回Excel/CSV文件，完成数据编辑闭环。接受parse_sheet返回的JSON格式数据（修改后）。保留原文件格式和样式，默认创建备份文件防止数据丢失。支持添加、删除、修改行和单元格数据。",
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

    try:
        result = core_service.convert_to_html(
            arguments["file_path"],
            arguments.get("output_path"),
            sheet_name=arguments.get("sheet_name"),
            page_size=arguments.get("page_size"),
            page_number=arguments.get("page_number"),
            header_rows=arguments.get("header_rows", 1)
        )

        # 结构化成功响应，便于LLM理解
        response = {
            "success": True,
            "operation": "convert_to_html",
            "results": result,
            "summary": {
                "files_generated": len(result),
                "total_size_kb": sum(r.get("file_size_kb", 0) for r in result),
                "sheets_converted": [r.get("sheet_name") for r in result]
            }
        }

        return [TextContent(
            type="text",
            text=json.dumps(response, ensure_ascii=False, indent=2)
        )]

    except FileNotFoundError as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "file_not_found",
                "error_message": f"文件未找到: {str(e)}",
                "suggestion": "请检查文件路径是否正确，确保文件存在且可访问。支持的格式: .xlsx, .xls, .xlsb, .xlsm, .csv"
            }, ensure_ascii=False, indent=2)
        )]
    except PermissionError as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "permission_error",
                "error_message": f"权限不足: {str(e)}",
                "suggestion": "请检查文件权限，确保有读取源文件和写入目标目录的权限"
            }, ensure_ascii=False, indent=2)
        )]
    except ValueError as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "invalid_parameter",
                "error_message": f"参数错误: {str(e)}",
                "suggestion": "请检查参数格式，如page_size和page_number应为正整数，sheet_name应为有效的工作表名称"
            }, ensure_ascii=False, indent=2)
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "conversion_error",
                "error_message": f"转换失败: {str(e)}",
                "suggestion": "请检查文件是否损坏，或尝试使用不同的参数。如果是大文件，建议使用page_size参数进行分页"
            }, ensure_ascii=False, indent=2)
        )]


async def _handle_parse_sheet(arguments: dict[str, Any], core_service: CoreService) -> list[TextContent]:
    """处理 parse_sheet 工具调用 - 优化版本，避免上下文爆炸。"""

    try:
        # 获取参数
        file_path = arguments["file_path"]
        sheet_name = arguments.get("sheet_name")
        range_string = arguments.get("range_string")
        include_full_data = arguments.get("include_full_data", False)
        include_styles = arguments.get("include_styles", False)
        preview_rows = arguments.get("preview_rows", 5)
        max_rows = arguments.get("max_rows")

        # 调用优化后的解析方法
        result = core_service.parse_sheet_optimized(
            file_path=file_path,
            sheet_name=sheet_name,
            range_string=range_string,
            include_full_data=include_full_data,
            include_styles=include_styles,
            preview_rows=preview_rows,
            max_rows=max_rows
        )

        # 为LLM添加使用指导
        result["llm_guidance"] = {
            "current_mode": "overview" if not include_full_data else "full_data",
            "next_steps": _generate_next_steps_guidance(result, include_full_data, include_styles),
            "data_access": {
                "headers": "result['headers'] - 列标题",
                "preview": "result['preview_rows'] - 数据预览",
                "full_data": "设置 include_full_data=true 获取完整数据",
                "styles": "设置 include_styles=true 获取样式信息"
            }
        }

        response = {
            "success": True,
            "operation": "parse_sheet",
            "data": result
        }

        return [TextContent(
            type="text",
            text=json.dumps(response, ensure_ascii=False, indent=2)
        )]

    except FileNotFoundError as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "file_not_found",
                "error_message": f"文件未找到: {str(e)}",
                "suggestion": "请检查文件路径是否正确。支持的格式: .xlsx, .xls, .xlsb, .xlsm, .csv"
            }, ensure_ascii=False, indent=2)
        )]
    except ValueError as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "invalid_parameter",
                "error_message": f"参数错误: {str(e)}",
                "suggestion": "请检查sheet_name是否存在，range_string格式是否正确（如'A1:D10'）"
            }, ensure_ascii=False, indent=2)
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "parsing_error",
                "error_message": f"解析失败: {str(e)}",
                "suggestion": "请检查文件是否损坏，或尝试指定具体的工作表名称和范围"
            }, ensure_ascii=False, indent=2)
        )]


def _generate_next_steps_guidance(result: dict, include_full_data: bool, include_styles: bool) -> list[str]:
    """生成下一步操作建议。"""
    guidance = []

    if not include_full_data:
        total_rows = result.get("metadata", {}).get("total_rows", 0)
        if total_rows > result.get("metadata", {}).get("preview_rows", 5):
            guidance.append(f"文件包含{total_rows}行数据，当前只显示预览。设置include_full_data=true获取完整数据")

    if not include_styles and result.get("metadata", {}).get("has_styles", False):
        guidance.append("文件包含样式信息（字体、颜色等）。设置include_styles=true获取样式数据")

    if result.get("metadata", {}).get("total_cells", 0) > 1000:
        guidance.append("文件较大，建议使用range_string参数获取特定范围，如'A1:D10'")

    if len(guidance) == 0:
        guidance.append("数据已完整加载，可以进行分析或修改")

    return guidance


async def _handle_apply_changes(arguments: dict[str, Any], core_service: CoreService) -> list[TextContent]:
    """处理 apply_changes 工具调用。"""

    try:
        result = core_service.apply_changes(
            arguments["file_path"],
            arguments["table_model_json"],
            arguments.get("create_backup", True)
        )

        response = {
            "success": True,
            "operation": "apply_changes",
            "result": result,
            "message": "数据已成功写回文件",
            "backup_created": arguments.get("create_backup", True)
        }

        return [TextContent(
            type="text",
            text=json.dumps(response, ensure_ascii=False, indent=2)
        )]

    except FileNotFoundError as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "file_not_found",
                "error_message": f"文件未找到: {str(e)}",
                "suggestion": "请检查文件路径是否正确，确保目标文件存在"
            }, ensure_ascii=False, indent=2)
        )]
    except PermissionError as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "permission_error",
                "error_message": f"权限不足: {str(e)}",
                "suggestion": "请检查文件权限，确保有写入文件的权限，文件可能被其他程序占用"
            }, ensure_ascii=False, indent=2)
        )]
    except ValueError as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "invalid_data",
                "error_message": f"数据格式错误: {str(e)}",
                "suggestion": "请确保table_model_json格式正确，包含必需的字段：sheet_name, headers, rows"
            }, ensure_ascii=False, indent=2)
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=json.dumps({
                "success": False,
                "error_type": "write_error",
                "error_message": f"写入失败: {str(e)}",
                "suggestion": "请检查数据格式是否与原文件兼容，或尝试关闭可能占用文件的程序"
            }, ensure_ascii=False, indent=2)
        )]
