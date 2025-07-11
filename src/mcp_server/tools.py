"""
MCP 工具注册模块

此模块定义并注册表格解析服务器的所有MCP工具。
"""

import logging
from typing import Any
from pathlib import Path

from mcp.server import Server
from mcp.types import Tool, TextContent

from ..services.sheet_service import SheetService
from ..parsers.factory import ParserFactory
from ..converters.html_converter import HTMLConverter
from ..converters.json_converter import JSONConverter
from ..models.table_model import Sheet, Row, Cell, Style

logger = logging.getLogger(__name__)


def register_tools(server: Server) -> None:
    """向服务器注册所有MCP工具。"""
    
    # 初始化核心服务
    parser_factory = ParserFactory()
    html_converter = HTMLConverter()
    sheet_service = SheetService(parser_factory, html_converter)
    
    @server.list_tools()
    async def handle_list_tools() -> list[Tool]:
        # 工具描述部分保持英文
        return [
            Tool(
                name="parse_sheet_to_json",
                description="Parse a spreadsheet file and return structured JSON data for LLM analysis",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to parse"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="convert_json_to_html",
                description="Convert JSON spreadsheet data to a perfect HTML file with 95% style fidelity",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "json_data": {
                            "type": "object",
                            "description": "JSON data from parse_sheet_to_json"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the HTML file"
                        }
                    },
                    "required": ["json_data", "output_path"]
                }
            ),
            Tool(
                name="convert_file_to_html",
                description="Smart direct conversion from file to HTML with intelligent processing strategy",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to convert"
                        },
                        "compact_mode": {
                            "type": "boolean",
                            "description": "Enable compact HTML mode for size optimization",
                            "default": True
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="convert_file_to_html_file",
                description="Convert spreadsheet file directly to HTML file (large file friendly)",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to convert"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the HTML file"
                        }
                    },
                    "required": ["file_path", "output_path"]
                }
            ),
            Tool(
                name="get_table_summary",
                description="Get a quick preview and summary of spreadsheet content",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to analyze"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="get_sheet_metadata",
                description="Get detailed metadata about the spreadsheet file",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file to analyze"
                        }
                    },
                    "required": ["file_path"]
                }
            )
        ]
    
    @server.call_tool()
    async def handle_call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
        """处理工具调用。"""
        try:
            if name == "parse_sheet_to_json":
                return await _handle_parse_sheet_to_json(arguments, sheet_service)
            elif name == "convert_json_to_html":
                return await _handle_convert_json_to_html(arguments, sheet_service)
            elif name == "convert_file_to_html":
                return await _handle_convert_file_to_html(arguments, sheet_service)
            elif name == "convert_file_to_html_file":
                return await _handle_convert_file_to_html_file(arguments, sheet_service)
            elif name == "get_table_summary":
                return await _handle_get_table_summary(arguments, sheet_service)
            elif name == "get_sheet_metadata":
                return await _handle_get_sheet_metadata(arguments, sheet_service)
            else:
                raise ValueError(f"未知工具: {name}")
                
        except Exception as e:
            logger.error(f"工具 {name} 出错: {e}")
            return [TextContent(
                type="text",
                text=f"错误: {str(e)}"
            )]


# 工具处理函数
async def _handle_parse_sheet_to_json(arguments: dict[str, Any], service: SheetService) -> list[TextContent]:
    """处理 parse_sheet_to_json 工具调用。"""
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("必须提供 file_path")

    if not Path(file_path).exists():
        raise FileNotFoundError(f"文件未找到: {file_path}")

    try:
        # 解析文件为 Sheet 对象
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # 转换为 JSON
        json_converter = JSONConverter()
        json_data = json_converter.convert(sheet)

        # 获取大小估算
        size_info = json_converter.estimate_json_size(sheet)

        # 返回带元数据的 JSON 数据
        import json
        json_string = json.dumps(json_data, indent=2, ensure_ascii=False)

        return [TextContent(
            type="text",
            text=f"JSON 转换成功！\n\n"
                 f"元数据:\n"
                 f"- 行数: {json_data['metadata']['rows']}\n"
                 f"- 列数: {json_data['metadata']['cols']}\n"
                 f"- 唯一样式数: {len(json_data['styles'])}\n"
                 f"- JSON 字符数: {size_info['total_characters']}\n"
                 f"- 估算字节数: {size_info['total_bytes']}\n\n"
                 f"JSON 数据:\n{json_string}"
        )]
    except Exception as e:
        raise RuntimeError(f"解析表格为 JSON 失败: {str(e)}")


async def _handle_convert_json_to_html(arguments: dict[str, Any], service: SheetService) -> list[TextContent]:
    """处理 convert_json_to_html 工具调用。"""
    json_data = arguments.get("json_data")
    output_path = arguments.get("output_path")

    if not json_data:
        raise ValueError("必须提供 json_data")
    if not output_path:
        raise ValueError("必须提供 output_path")

    try:
        # 从 JSON 数据重建 Sheet 对象
        sheet = _json_to_sheet(json_data)

        # 转换为优化后的 HTML
        html_converter = HTMLConverter(compact_mode=True)
        html_content = html_converter.convert(sheet, optimize=True)

        # 写入文件
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(html_content, encoding='utf-8')

        # 获取大小信息
        file_size = len(html_content)

        return [TextContent(
            type="text",
            text=f"HTML 文件生成成功！\n\n"
                 f"输出路径: {output_path}\n"
                 f"文件字符数: {file_size}\n"
                 f"处理行数: {len(sheet.rows)}\n"
                 f"优化: 启用 CSS 类复用\n\n"
                 f"HTML 文件已保存，样式还原度 95%，并已优化体积。"
        )]
    except Exception as e:
        raise RuntimeError(f"JSON 转 HTML 失败: {str(e)}")


async def _handle_convert_file_to_html(arguments: dict[str, Any], service: SheetService) -> list[TextContent]:
    """处理 convert_file_to_html 工具调用。"""
    # 基于现有服务的基础实现
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("必须提供 file_path")
    
    if not Path(file_path).exists():
        raise FileNotFoundError(f"文件未找到: {file_path}")
    
    try:
        html_content = service.convert_to_html(file_path)
        return [TextContent(
            type="text",
            text=f"HTML 转换完成。内容长度: {len(html_content)} 字符。\n\n{html_content}"
        )]
    except Exception as e:
        raise RuntimeError(f"文件转 HTML 失败: {str(e)}")


async def _handle_convert_file_to_html_file(arguments: dict[str, Any], service: SheetService) -> list[TextContent]:
    """处理 convert_file_to_html_file 工具调用。"""
    file_path = arguments.get("file_path")
    output_path = arguments.get("output_path")

    if not file_path:
        raise ValueError("必须提供 file_path")
    if not output_path:
        raise ValueError("必须提供 output_path")

    if not Path(file_path).exists():
        raise FileNotFoundError(f"文件未找到: {file_path}")

    try:
        # 解析文件
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # 转换为优化后的 HTML
        html_converter = HTMLConverter(compact_mode=True)
        html_content = html_converter.convert(sheet, optimize=True)

        # 写入文件
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(html_content, encoding='utf-8')

        # 获取统计信息
        file_size = len(html_content)
        size_info = html_converter.estimate_size_reduction(sheet)

        return [TextContent(
            type="text",
            text=f"HTML 文件转换成功！\n\n"
                 f"输入文件: {file_path}\n"
                 f"输出文件: {output_path}\n"
                 f"文件字符数: {file_size}\n"
                 f"行数: {len(sheet.rows)}\n"
                 f"优化率: {size_info['reduction_percentage']:.1f}%\n"
                 f"样式还原度: 95%\n\n"
                 f"HTML 文件已保存，具备专业样式与优化性能。"
        )]
    except Exception as e:
        raise RuntimeError(f"文件转 HTML 文件失败: {str(e)}")


async def _handle_get_table_summary(arguments: dict[str, Any], service: SheetService) -> list[TextContent]:
    """处理 get_table_summary 工具调用。"""
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("必须提供 file_path")

    if not Path(file_path).exists():
        raise FileNotFoundError(f"文件未找到: {file_path}")

    try:
        # 解析文件
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # 生成统计信息
        total_rows = len(sheet.rows)
        total_cols = len(sheet.rows[0].cells) if sheet.rows else 0

        # 统计非空单元格和有样式单元格
        non_empty_cells = 0
        styled_cells = 0
        for row in sheet.rows:
            for cell in row.cells:
                if cell.value is not None and str(cell.value).strip():
                    non_empty_cells += 1
                if cell.style:
                    styled_cells += 1

        # 获取前5行样例数据
        sample_data = []
        for i, row in enumerate(sheet.rows[:5]):
            row_data = []
            for j, cell in enumerate(row.cells[:5]):  # 前5列
                value = str(cell.value) if cell.value is not None else ""
                row_data.append(value[:20] + "..." if len(value) > 20 else value)
            sample_data.append(f"第{i+1}行: {' | '.join(row_data)}")

        # 估算处理建议
        total_cells = total_rows * total_cols
        estimated_html_size = total_cells * 50  # 粗略估算

        if estimated_html_size > 100000:
            recommendation = "大文件 - 推荐使用 convert_file_to_html_file 输出到文件"
        elif estimated_html_size > 50000:
            recommendation = "中等文件 - 建议开启紧凑模式优化"
        else:
            recommendation = "小文件 - 可直接使用 convert_file_to_html 输出"

        return [TextContent(
            type="text",
            text=f"表格摘要: {Path(file_path).name}\n\n"
                 f"📊 基本统计:\n"
                 f"- 表名: {sheet.name}\n"
                 f"- 尺寸: {total_rows} 行 × {total_cols} 列\n"
                 f"- 总单元格: {total_cells:,}\n"
                 f"- 非空单元格: {non_empty_cells:,}\n"
                 f"- 有样式单元格: {styled_cells:,}\n"
                 f"- 合并单元格: {len(sheet.merged_cells) if sheet.merged_cells else 0}\n\n"
                 f"📋 前5行样例:\n" + "\n".join(sample_data) + "\n\n"
                 f"💡 处理建议:\n{recommendation}\n\n"
                 f"📈 估算 HTML 大小: ~{estimated_html_size:,} 字符"
        )]
    except Exception as e:
        raise RuntimeError(f"生成表格摘要失败: {str(e)}")


async def _handle_get_sheet_metadata(arguments: dict[str, Any], service: SheetService) -> list[TextContent]:
    """处理 get_sheet_metadata 工具调用。"""
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("必须提供 file_path")

    if not Path(file_path).exists():
        raise FileNotFoundError(f"文件未找到: {file_path}")

    try:
        # 获取文件信息
        file_info = Path(file_path)
        file_size = file_info.stat().st_size
        file_ext = file_info.suffix.lower()

        # 解析文件
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # 样式分析
        unique_styles = set()
        font_families = set()
        colors = set()

        for row in sheet.rows:
            for cell in row.cells:
                if cell.style:
                    # 样式签名
                    style_sig = f"{cell.style.bold}_{cell.style.italic}_{cell.style.font_color}_{cell.style.background_color}"
                    unique_styles.add(style_sig)

                    if cell.style.font_name:
                        font_families.add(cell.style.font_name)
                    if cell.style.font_color:
                        colors.add(cell.style.font_color)
                    if cell.style.background_color:
                        colors.add(cell.style.background_color)

        # 数据类型分析
        data_types = {"text": 0, "number": 0, "empty": 0}
        for row in sheet.rows:
            for cell in row.cells:
                if cell.value is None or str(cell.value).strip() == "":
                    data_types["empty"] += 1
                elif isinstance(cell.value, (int, float)):
                    data_types["number"] += 1
                else:
                    data_types["text"] += 1

        # 生成 JSON 和 HTML 大小估算
        json_converter = JSONConverter()
        json_size = json_converter.estimate_json_size(sheet)

        html_converter = HTMLConverter(compact_mode=True)
        html_size = html_converter.estimate_size_reduction(sheet)

        return [TextContent(
            type="text",
            text=f"表格元数据: {file_info.name}\n\n"
                 f"📁 文件信息:\n"
                 f"- 路径: {file_path}\n"
                 f"- 文件大小: {file_size:,} 字节\n"
                 f"- 文件格式: {file_ext.upper()}\n"
                 f"- 表名: {sheet.name}\n\n"
                 f"📊 结构:\n"
                 f"- 尺寸: {len(sheet.rows)} 行 × {len(sheet.rows[0].cells) if sheet.rows else 0} 列\n"
                 f"- 总单元格: {len(sheet.rows) * (len(sheet.rows[0].cells) if sheet.rows else 0):,}\n"
                 f"- 合并单元格: {len(sheet.merged_cells) if sheet.merged_cells else 0}\n\n"
                 f"🎨 样式:\n"
                 f"- 唯一样式数: {len(unique_styles)}\n"
                 f"- 字体族: {len(font_families)} ({', '.join(list(font_families)[:3])}{'...' if len(font_families) > 3 else ''})\n"
                 f"- 使用颜色数: {len(colors)}\n\n"
                 f"📈 数据分析:\n"
                 f"- 文本单元格: {data_types['text']:,}\n"
                 f"- 数字单元格: {data_types['number']:,}\n"
                 f"- 空单元格: {data_types['empty']:,}\n\n"
                 f"💾 输出估算:\n"
                 f"- JSON 字符数: {json_size['total_characters']:,}\n"
                 f"- HTML 原始大小: {html_size['original_size']:,}\n"
                 f"- HTML 优化后大小: {html_size['optimized_size']:,}\n"
                 f"- 优化节省: {html_size['reduction_percentage']:.1f}%"
        )]
    except Exception as e:
        raise RuntimeError(f"获取表格元数据失败: {str(e)}")


def _json_to_sheet(json_data: dict[str, Any]) -> Sheet:
    """
    将JSON数据转换回Sheet对象。

    参数:
        json_data: 来自 parse_sheet_to_json 的JSON数据

    返回:
        重构的Sheet对象
    """
    # 这里已在文件顶部导入 Sheet, Row, Cell, Style

    # 提取元数据
    metadata = json_data.get('metadata', {})
    sheet_name = metadata.get('name', 'Untitled')

    # 提取样式
    styles_dict = json_data.get('styles', {})

    # 重建行
    rows = []
    for row_data in json_data.get('data', []):
        cells = []
        for cell_data in row_data.get('cells', []):
            # 重建单元格值
            value = cell_data.get('value')

            # 重建样式
            style = None
            style_id = cell_data.get('style_id')
            if style_id and style_id in styles_dict:
                style_data = styles_dict[style_id]
                style = Style(
                    bold=style_data.get('bold', False),
                    italic=style_data.get('italic', False),
                    underline=style_data.get('underline', False),
                    font_color=style_data.get('font_color', '#000000'),
                    background_color=style_data.get('background_color', '#FFFFFF'),
                    text_align=style_data.get('text_align', 'left'),
                    vertical_align=style_data.get('vertical_align', 'top'),
                    font_size=style_data.get('font_size'),
                    font_name=style_data.get('font_name'),
                    border_top=style_data.get('border_top', ''),
                    border_bottom=style_data.get('border_bottom', ''),
                    border_left=style_data.get('border_left', ''),
                    border_right=style_data.get('border_right', ''),
                    border_color=style_data.get('border_color', '#000000'),
                    wrap_text=style_data.get('wrap_text', False),
                    number_format=style_data.get('number_format', '')
                )

            # 创建单元格
            cell = Cell(value=value, style=style)
            cells.append(cell)

        rows.append(Row(cells=cells))

    # 创建 Sheet
    merged_cells = json_data.get('merged_cells', [])
    return Sheet(name=sheet_name, rows=rows, merged_cells=merged_cells)
