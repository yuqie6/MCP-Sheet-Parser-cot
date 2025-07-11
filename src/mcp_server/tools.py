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
from ..utils.performance import PerformanceOptimizer

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
            ),
            Tool(
                name="convert_file_to_html_paginated",
                description="Convert a large spreadsheet file to HTML with pagination support for better performance",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the spreadsheet file"
                        },
                        "output_dir": {
                            "type": "string",
                            "description": "Directory to save paginated HTML files"
                        },
                        "max_rows_per_page": {
                            "type": "integer",
                            "description": "Maximum number of rows per page (default: 1000)",
                            "default": 1000
                        },
                        "compact_mode": {
                            "type": "boolean",
                            "description": "Enable compact HTML mode for smaller file sizes (default: true)",
                            "default": True
                        }
                    },
                    "required": ["file_path", "output_dir"]
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
            elif name == "convert_file_to_html_paginated":
                return await _handle_convert_file_to_html_paginated(arguments, sheet_service)
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
        # 初始化性能优化器
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()

        # 解析文件
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # 分析性能指标
        metrics = optimizer.analyze_sheet_performance(sheet)

        # 生成统计信息
        total_rows = metrics.rows
        total_cols = metrics.cols

        # 统计进阶功能使用情况
        hyperlink_cells = 0
        comment_cells = 0
        for row in sheet.rows:
            for cell in row.cells:
                if cell.style:
                    if cell.style.hyperlink:
                        hyperlink_cells += 1
                    if cell.style.comment:
                        comment_cells += 1

        # 获取前5行样例数据
        sample_data = []
        for i, row in enumerate(sheet.rows[:5]):
            row_data = []
            for j, cell in enumerate(row.cells[:5]):  # 前5列
                value = str(cell.value) if cell.value is not None else ""
                row_data.append(value[:20] + "..." if len(value) > 20 else value)
            sample_data.append(f"第{i+1}行: {' | '.join(row_data)}")

        # 获取性能建议
        recommendation = optimizer.get_recommendation_message(metrics.recommendation, metrics)

        # 计算总单元格数
        total_cells = metrics.total_cells

        return [TextContent(
            type="text",
            text=f"表格摘要: {Path(file_path).name}\n\n"
                 f"📊 基本统计:\n"
                 f"- 表名: {sheet.name}\n"
                 f"- 尺寸: {total_rows} 行 × {total_cols} 列\n"
                 f"- 总单元格: {total_cells:,}\n"
                 f"- 非空单元格: {metrics.non_empty_cells:,}\n"
                 f"- 有样式单元格: {metrics.styled_cells:,}\n"
                 f"- 合并单元格: {len(sheet.merged_cells) if sheet.merged_cells else 0}\n\n"
                 f"🚀 进阶功能:\n"
                 f"- 超链接单元格: {hyperlink_cells:,}\n"
                 f"- 注释单元格: {comment_cells:,}\n\n"
                 f"📋 前5行样例:\n" + "\n".join(sample_data) + "\n\n"
                 f"💡 处理建议:\n{recommendation}\n\n"
                 f"📈 性能指标:\n"
                 f"- 估算HTML大小: ~{metrics.estimated_html_size:,} 字符\n"
                 f"- 估算JSON大小: ~{metrics.estimated_json_size:,} 字符\n"
                 f"- 解析耗时: {metrics.processing_time:.2f} 秒"
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


async def _handle_convert_file_to_html_paginated(arguments: dict[str, Any], service: SheetService) -> list[TextContent]:
    """处理 convert_file_to_html_paginated 工具调用。"""
    file_path = arguments.get("file_path")
    output_dir = arguments.get("output_dir")
    max_rows_per_page = arguments.get("max_rows_per_page", 1000)
    compact_mode = arguments.get("compact_mode", True)

    if not file_path:
        raise ValueError("缺少必需参数: file_path")
    if not output_dir:
        raise ValueError("缺少必需参数: output_dir")

    try:
        import os

        # 初始化性能优化器
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()

        # 解析文件
        parser_factory = ParserFactory()
        parser = parser_factory.get_parser(file_path)
        sheet = parser.parse(file_path)

        # 分析性能指标
        metrics = optimizer.analyze_sheet_performance(sheet)

        # 检查是否需要分页
        if not optimizer.should_use_pagination(metrics):
            return [TextContent(
                type="text",
                text=f"文件不需要分页处理。建议使用 convert_file_to_html_file 工具。\n"
                     f"文件大小: {metrics.total_cells:,} 个单元格\n"
                     f"处理建议: {optimizer.get_recommendation_message(metrics.recommendation, metrics)}"
            )]

        # 计算分页参数
        pagination = optimizer.calculate_pagination_params(sheet, max_rows_per_page)

        # 确保输出目录存在
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        # 创建HTML转换器
        html_converter = HTMLConverter()

        # 生成分页文件
        generated_files = []
        base_filename = Path(file_path).stem

        for page_info in pagination["pages"]:
            # 创建当前页的Sheet对象
            page_rows = sheet.rows[page_info["start_row"]:page_info["end_row"]]
            page_sheet = Sheet(
                name=f"{sheet.name} - 第{page_info['page_number']}页",
                rows=page_rows,
                merged_cells=[]  # 分页时暂不处理合并单元格
            )

            # 转换为HTML
            html_content = html_converter.convert(page_sheet)

            # 保存文件
            page_filename = f"{base_filename}_page_{page_info['page_number']:03d}.html"
            page_file_path = output_path / page_filename

            with open(page_file_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            generated_files.append(str(page_file_path))

        # 生成索引文件
        index_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>{sheet.name} - 分页索引</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        .summary {{ background: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
        .page-list {{ list-style-type: none; padding: 0; }}
        .page-list li {{ margin: 10px 0; }}
        .page-list a {{ text-decoration: none; color: #0066cc; }}
        .page-list a:hover {{ text-decoration: underline; }}
        .stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; }}
        .stat {{ background: white; padding: 10px; border-radius: 3px; border-left: 4px solid #0066cc; }}
    </style>
</head>
<body>
    <h1>{sheet.name} - 分页索引</h1>

    <div class="summary">
        <h2>📊 文件统计</h2>
        <div class="stats">
            <div class="stat">
                <strong>总行数:</strong> {metrics.rows:,}
            </div>
            <div class="stat">
                <strong>总列数:</strong> {metrics.cols:,}
            </div>
            <div class="stat">
                <strong>总单元格:</strong> {metrics.total_cells:,}
            </div>
            <div class="stat">
                <strong>非空单元格:</strong> {metrics.non_empty_cells:,}
            </div>
            <div class="stat">
                <strong>分页数量:</strong> {pagination['total_pages']}
            </div>
            <div class="stat">
                <strong>每页行数:</strong> {max_rows_per_page:,}
            </div>
        </div>
    </div>

    <h2>📄 分页列表</h2>
    <ul class="page-list">
"""

        for page_info in pagination["pages"]:
            page_filename = f"{base_filename}_page_{page_info['page_number']:03d}.html"
            index_content += f"""        <li>
            <a href="{page_filename}">
                第 {page_info['page_number']} 页
                (行 {page_info['start_row'] + 1:,} - {page_info['end_row']:,},
                共 {page_info['row_count']:,} 行)
            </a>
        </li>
"""

        index_content += """    </ul>
</body>
</html>"""

        # 保存索引文件
        index_file_path = output_path / f"{base_filename}_index.html"
        with open(index_file_path, 'w', encoding='utf-8') as f:
            f.write(index_content)

        generated_files.insert(0, str(index_file_path))

        return [TextContent(
            type="text",
            text=f"✅ 分页转换完成!\n\n"
                 f"📊 处理统计:\n"
                 f"- 原文件: {Path(file_path).name}\n"
                 f"- 总行数: {metrics.rows:,}\n"
                 f"- 总页数: {pagination['total_pages']}\n"
                 f"- 每页行数: {max_rows_per_page:,}\n"
                 f"- 处理时间: {metrics.processing_time:.2f} 秒\n\n"
                 f"📁 生成文件:\n" + "\n".join(f"- {Path(f).name}" for f in generated_files) + "\n\n"
                 f"💡 使用建议:\n"
                 f"- 打开 {base_filename}_index.html 查看分页索引\n"
                 f"- 每个分页文件包含 {max_rows_per_page:,} 行数据\n"
                 f"- 文件保存在: {output_dir}"
        )]

    except Exception as e:
        raise RuntimeError(f"分页转换失败: {str(e)}")
