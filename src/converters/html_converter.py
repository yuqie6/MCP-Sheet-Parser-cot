"""
HTML转换器模块

将Sheet对象转换为高保真度的HTML文件
"""

import datetime
import logging
import re
from pathlib import Path
from typing import Any
from src.models.table_model import Sheet, Cell, Style
from src.utils.range_parser import parse_range_string
from src.constants import StyleConstants, HTMLConstants, Limits
from src.exceptions import HTMLConversionError, ValidationError
from src.font_manager import get_font_manager

logger = logging.getLogger(__name__)

# 边框样式映射常量
BORDER_STYLE_MAP = {
    "thin": ("1px", "solid"),
    "medium": ("2px", "solid"),
    "thick": ("3px", "solid"),
    "solid": ("1px", "solid"),
    "dashed": ("1px", "dashed"),
    "dotted": ("1px", "dotted"),
    "double": ("3px", "double"),
    "groove": ("2px", "groove"),
    "ridge": ("2px", "ridge"),
    "inset": ("2px", "inset"),
    "outset": ("2px", "outset"),
    "hair": ("1px", "solid"),
    "mediumdashed": ("2px", "dashed"),
    "dashdot": ("1px", "dashed"),
    "mediumdashdot": ("2px", "dashed"),
    "dashdotdot": ("1px", "dashed"),
    "mediumdashdotdot": ("2px", "dashed"),
    "slantdashdot": ("1px", "dashed")
}

# 数字格式映射常量
NUMBER_FORMAT_MAP = {
    "General": lambda v: str(v),
    "0": lambda v: f"{v:.0f}" if isinstance(v, (int, float)) else str(v),
    "0.0": lambda v: f"{v:.1f}" if isinstance(v, (int, float)) else str(v),
    "0.00": lambda v: f"{v:.2f}" if isinstance(v, (int, float)) else str(v),
    "#,##0": lambda v: f"{v:,.0f}" if isinstance(v, (int, float)) else str(v),
    "#,##0.0": lambda v: f"{v:,.1f}" if isinstance(v, (int, float)) else str(v),
    "#,##0.00": lambda v: f"{v:,.2f}" if isinstance(v, (int, float)) else str(v),
    "0%": lambda v: f"{v:.0%}" if isinstance(v, (int, float)) else str(v),
    "0.0%": lambda v: f"{v:.1%}" if isinstance(v, (int, float)) else str(v),
    "0.00%": lambda v: f"{v:.2%}" if isinstance(v, (int, float)) else str(v),
}

# 日期格式映射常量
DATE_FORMAT_MAP = {
    "yyyy-mm-dd": "%Y-%m-%d",
    "mm/dd/yyyy": "%m/%d/%Y",
    "dd/mm/yyyy": "%d/%m/%Y",
    "yyyy/mm/dd": "%Y/%m/%d",
    "mm-dd-yyyy": "%m-%d-%Y",
    "dd-mm-yyyy": "%d-%m-%Y",
}



class HTMLConverter:
    """HTML转换器，将Sheet对象转换为完美的HTML文件。"""
    
    def __init__(self, compact_mode: bool = False, header_rows: int = 1):
        """
        初始化HTML转换器。

        Args:
            compact_mode: 是否使用紧凑模式（减少空白字符）
            header_rows: 表头行数，默认第一行为表头
        """
        self.compact_mode = compact_mode
        self.header_rows = header_rows
    
    def convert_to_file(self, sheet: Sheet, output_path: str) -> dict[str, Any]:
        """
        将Sheet对象转换为HTML文件。
        
        Args:
            sheet: Sheet对象
            output_path: 输出文件路径
            
        Returns:
            转换结果信息
        """
        try:
            # 生成HTML内容
            html_content = self._generate_html(sheet)
            
            # 写入文件
            output_file = Path(output_path)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            # 计算文件大小
            file_size = output_file.stat().st_size
            
            return {
                "status": "success",
                "output_path": str(output_file.absolute()),
                "file_size": file_size,
                "file_size_kb": round(file_size / 1024, 2),
                "sheet_name": sheet.name,
                "rows_converted": len(sheet.rows),
                "cells_converted": sum(len(row.cells) for row in sheet.rows),
                "has_styles": any(any(cell.style for cell in row.cells) for row in sheet.rows),
                "has_merged_cells": len(sheet.merged_cells) > 0
            }
            
        except Exception as e:
            logger.error(f"HTML转换失败: {e}")
            raise RuntimeError(f"HTML转换失败: {str(e)}")
    
    def _generate_html(self, sheet: Sheet) -> str:
        """
        生成完整的HTML内容。
        
        Args:
            sheet: Sheet对象
            
        Returns:
            HTML字符串
        """
        # 收集所有样式
        styles = self._collect_styles(sheet)
        
        # 生成CSS
        css_content = self._generate_css(styles)
        
        # 生成表格HTML
        table_html = self._generate_table(sheet, styles)
        
        # 组装完整HTML
        html_template = self._get_html_template()
        
        html_content = html_template.format(
            title=f"表格: {sheet.name}",
            css_styles=css_content,
            table_content=table_html,
            sheet_name=sheet.name
        )
        
        if self.compact_mode:
            html_content = self._compact_html(html_content)
        
        return html_content
    
    def _collect_styles(self, sheet: Sheet) -> dict[str, Style]:
        """
        收集所有唯一的样式。
        
        Args:
            sheet: Sheet对象
            
        Returns:
            样式字典，键为样式ID，值为Style对象
        """
        styles = {}
        style_counter = 0
        
        for row in sheet.rows:
            for cell in row.cells:
                if cell.style:
                    # 生成样式的唯一标识
                    style_key = self._get_style_key(cell.style)
                    if style_key not in styles:
                        style_counter += 1
                        styles[f"style_{style_counter}"] = cell.style
        
        return styles
    
    def _get_style_key(self, style: Style) -> str:
        """
        生成样式的唯一标识。
        
        Args:
            style: Style对象
            
        Returns:
            样式的唯一标识字符串
        """
        # 基于样式属性生成唯一键
        key_parts = []
        
        if style.font_name:
            key_parts.append(f"fn:{style.font_name}")
        if style.font_size:
            key_parts.append(f"fs:{style.font_size}")
        if style.font_color:
            key_parts.append(f"fc:{style.font_color}")
        if style.background_color:
            key_parts.append(f"bg:{style.background_color}")
        if style.bold:
            key_parts.append("bold")
        if style.italic:
            key_parts.append("italic")
        if style.underline:
            key_parts.append("underline")
        if style.text_align:
            key_parts.append(f"ta:{style.text_align}")
        if style.vertical_align:
            key_parts.append(f"va:{style.vertical_align}")

        # 边框属性
        if style.border_top:
            key_parts.append(f"bt:{style.border_top}")
        if style.border_bottom:
            key_parts.append(f"bb:{style.border_bottom}")
        if style.border_left:
            key_parts.append(f"bl:{style.border_left}")
        if style.border_right:
            key_parts.append(f"br:{style.border_right}")
        if style.border_color:
            key_parts.append(f"bc:{style.border_color}")

        # 文本换行和格式化
        if style.wrap_text:
            key_parts.append("wrap")
        if style.number_format:
            key_parts.append(f"nf:{style.number_format}")

        # 进阶功能
        if style.hyperlink:
            key_parts.append(f"link:{style.hyperlink}")
        if style.comment:
            key_parts.append(f"comment:{style.comment}")

        return "|".join(key_parts) if key_parts else "default"
    
    def _generate_css(self, styles: dict[str, Style]) -> str:
        """
        生成CSS样式。
        
        Args:
            styles: 样式字典
            
        Returns:
            CSS字符串
        """
        css_rules = []
        
        # 基础表格样式
        css_rules.append("""
        table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
            color: #000000 !important;  /* 强制默认文字颜色为黑色 */
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
            color: #000000 !important;  /* 强制确保默认文字颜色为黑色 */
        }
        th {
            background-color: #f2f2f2;
            font-weight: bold;
            color: #000000 !important;  /* 强制表头文字为黑色 */
        }
        /* 覆盖可能的暗色主题影响 */
        body {
            color: #000000 !important;
            background-color: #ffffff !important;
        }
        """)
        
        # 生成样式类
        for style_id, style in styles.items():
            css_rule = f".{style_id} {{"

            # 字体属性
            if style.font_name:
                # 智能处理字体名称，添加引号和回退字体
                font_family = self._format_font_family(style.font_name)
                css_rule += f" font-family: {font_family};"
            if style.font_size:
                # 优化字体大小单位转换
                font_size = self._format_font_size(style.font_size)
                css_rule += f" font-size: {font_size};"
            if style.font_color:
                # 验证和格式化颜色
                formatted_color = self._format_color(style.font_color)
                css_rule += f" color: {formatted_color} !important;"
            if style.bold:
                css_rule += " font-weight: bold;"
            if style.italic:
                css_rule += " font-style: italic;"
            if style.underline:
                css_rule += " text-decoration: underline;"

            # 背景和填充
            if style.background_color:
                formatted_bg_color = self._format_color(style.background_color)
                css_rule += f" background-color: {formatted_bg_color};"

            # 文本对齐
            if style.text_align:
                css_rule += f" text-align: {style.text_align};"
            if style.vertical_align:
                css_rule += f" vertical-align: {style.vertical_align};"

            # 边框属性
            border_styles = self._generate_border_css(style)
            if border_styles:
                css_rule += border_styles

            # 文本换行
            if style.wrap_text:
                css_rule += " white-space: pre-wrap; word-wrap: break-word;"

            # 数字格式（作为data属性，便于JavaScript处理）
            if style.number_format:
                # 将数字格式信息添加为注释，便于调试
                css_rule += f" /* number-format: {style.number_format} */"

            css_rule += " }"
            css_rules.append(css_rule)
        
        return "\n".join(css_rules)

    def _generate_border_css(self, style: Style) -> str:
        """
        生成边框CSS样式。

        Args:
            style: Style对象

        Returns:
            边框CSS字符串
        """
        border_css = ""

        # 边框颜色
        border_color = style.border_color if style.border_color else "#000000"

        # 处理各个边框
        if style.border_top:
            border_style_css = self._parse_border_style_complete(style.border_top, border_color)
            border_css += f" border-top: {border_style_css};"

        if style.border_bottom:
            border_style_css = self._parse_border_style_complete(style.border_bottom, border_color)
            border_css += f" border-bottom: {border_style_css};"

        if style.border_left:
            border_style_css = self._parse_border_style_complete(style.border_left, border_color)
            border_css += f" border-left: {border_style_css};"

        if style.border_right:
            border_style_css = self._parse_border_style_complete(style.border_right, border_color)
            border_css += f" border-right: {border_style_css};"

        return border_css



    def _parse_border_style_complete(self, border_style: str, border_color: str) -> str:
        """
        解析完整的边框样式，包括宽度、样式和颜色。

        Args:
            border_style: 边框样式字符串
            border_color: 边框颜色

        Returns:
            完整的CSS边框样式
        """
        if not border_style:
            return f"1px solid {border_color}"

        # 检查是否为已知样式
        border_lower = border_style.lower()
        if border_lower in BORDER_STYLE_MAP:
            width, style_type = BORDER_STYLE_MAP[border_lower]
            return f"{width} {style_type} {border_color}"

        # 尝试从字符串中提取数字和样式
        # 匹配数字+单位+样式的模式
        pattern = r'(\d+(?:\.\d+)?)(px|pt|em|rem)?\s*(solid|dashed|dotted|double|groove|ridge|inset|outset)?'
        match = re.search(pattern, border_style.lower())

        if match:
            width = match.group(1)
            unit = match.group(2) or "px"
            style_type = match.group(3) or "solid"
            return f"{width}{unit} {style_type} {border_color}"

        # 默认返回简单边框
        return f"1px solid {border_color}"

    def _format_font_family(self, font_name: str) -> str:
        """
        使用智能字体管理器格式化字体名称。

        Args:
            font_name: 原始字体名称

        Returns:
            格式化后的字体族字符串
        """
        font_manager = get_font_manager()
        return font_manager.generate_font_family(font_name)



    def _format_font_size(self, font_size: float) -> str:
        """
        格式化字体大小，优化单位转换。

        Args:
            font_size: 原始字体大小（通常为pt单位）

        Returns:
            格式化后的字体大小字符串
        """
        if not font_size or font_size <= 0:
            return f"{StyleConstants.DEFAULT_FONT_SIZE_PT}pt"  # 默认字体大小

        # Excel字体大小通常以pt为单位，但在Web中可能需要调整
        # 提供多种单位选择以获得更好的显示效果

        # 对于非常小或非常大的字体，进行适当调整
        if font_size < StyleConstants.MIN_FONT_SIZE_PT:
            adjusted_size = max(StyleConstants.MIN_FONT_SIZE_PT, font_size)  # 最小字体
        elif font_size > StyleConstants.MAX_FONT_SIZE_PT:
            adjusted_size = min(StyleConstants.MAX_FONT_SIZE_PT, font_size)  # 最大字体
        else:
            adjusted_size = font_size

        # 使用pt单位，因为它与Excel更一致
        pt_size = round(adjusted_size, 1)

        # 如果是整数，不显示小数点
        if pt_size == int(pt_size):
            return f"{int(pt_size)}pt"
        else:
            return f"{pt_size}pt"

    def _format_color(self, color: str) -> str:
        """
        验证和格式化颜色值。

        Args:
            color: 原始颜色值

        Returns:
            格式化后的颜色值
        """
        if not color:
            return StyleConstants.DEFAULT_FONT_COLOR  # 默认黑色

        # 移除空格和转换为大写
        color = color.strip().upper()

        # 如果已经是正确的十六进制格式
        if re.match(r'^#[0-9A-F]{6}$', color):
            return color

        # 如果是3位十六进制，扩展为6位
        if re.match(r'^#[0-9A-F]{3}$', color):
            return f"#{color[1]}{color[1]}{color[2]}{color[2]}{color[3]}{color[3]}"

        # 如果没有#前缀但是有效的十六进制
        if re.match(r'^[0-9A-F]{6}$', color):
            return f"#{color}"

        if re.match(r'^[0-9A-F]{3}$', color):
            return f"#{color[0]}{color[0]}{color[1]}{color[1]}{color[2]}{color[2]}"

        # 处理常见的颜色名称
        if color in StyleConstants.COLOR_NAMES:
            return StyleConstants.COLOR_NAMES[color]

        # 如果都不匹配，返回默认黑色
        return StyleConstants.DEFAULT_FONT_COLOR

    def _generate_table(self, sheet: Sheet, styles: dict[str, Style]) -> str:
        """
        生成表格HTML。
        
        Args:
            sheet: Sheet对象
            styles: 样式字典
            
        Returns:
            表格HTML字符串
        """
        
        # 处理合并单元格
        occupied_cells: set[tuple[int, int]] = set()
        merged_cells_map: dict[tuple[int, int], dict[str, int]] = {}

        for merged_range in sheet.merged_cells:
            try:
                start_row, start_col, end_row, end_col = parse_range_string(merged_range)
                row_span = end_row - start_row + 1
                col_span = end_col - start_col + 1
                
                if row_span > 1 or col_span > 1:
                    merged_cells_map[(start_row, start_col)] = {"rowspan": row_span, "colspan": col_span}
                    for r in range(start_row, end_row + 1):
                        for c in range(start_col, end_col + 1):
                            if (r, c) != (start_row, start_col):
                                occupied_cells.add((r, c))
            except ValueError as e:
                logger.warning(f"无法解析合并单元格范围 '{merged_range}': {e}")

        # 添加表格标题和描述
        table_parts = [f'<table role="table" aria-label="表格: {sheet.name}">']

        # 添加表格标题
        if sheet.name and sheet.name.strip():
            table_parts.append(f'<caption>表格: {self._escape_html(sheet.name)}</caption>')

        # Pre-compute a reverse map from style_key to style_id for efficiency
        style_key_to_id_map = {self._get_style_key(style_obj): style_id for style_id, style_obj in styles.items()}

        # 分离表头和表体
        if self.header_rows > 0 and len(sheet.rows) > 0:
            # 添加表头
            table_parts.append('<thead>')
            self._generate_rows_html(
                table_parts, sheet.rows[:self.header_rows],
                occupied_cells, merged_cells_map, style_key_to_id_map,
                is_header=True
            )
            table_parts.append('</thead>')

            # 添加表体
            if len(sheet.rows) > self.header_rows:
                table_parts.append('<tbody>')
                self._generate_rows_html(
                    table_parts, sheet.rows[self.header_rows:],
                    occupied_cells, merged_cells_map, style_key_to_id_map,
                    is_header=False, row_offset=self.header_rows
                )
                table_parts.append('</tbody>')
        else:
            # 没有表头，全部作为数据行
            self._generate_rows_html(
                table_parts, sheet.rows,
                occupied_cells, merged_cells_map, style_key_to_id_map,
                is_header=False
            )
        
        table_parts.append('</table>')
        
        return "\n".join(table_parts)

    def _generate_rows_html(self, table_parts: list, rows: list, occupied_cells: set,
                           merged_cells_map: dict, style_key_to_id_map: dict,
                           is_header: bool = False, row_offset: int = 0):
        """
        生成行HTML的通用方法，避免代码重复。

        Args:
            table_parts: HTML部分列表
            rows: 行列表
            occupied_cells: 被占用的单元格集合
            merged_cells_map: 合并单元格映射
            style_key_to_id_map: 样式键到ID的映射
            is_header: 是否为表头行
            row_offset: 行偏移量
        """
        for r_idx, row in enumerate(rows):
            actual_row_idx = r_idx + row_offset
            table_parts.append('<tr>')

            for c_idx, cell in enumerate(row.cells):
                if (actual_row_idx, c_idx) in occupied_cells:
                    continue

                # 确定样式类
                style_class = ""
                if cell.style:
                    style_key = self._get_style_key(cell.style)
                    style_id = style_key_to_id_map.get(style_key)
                    if style_id:
                        style_class = f' class="{style_id}"'

                span_attrs = ""
                if (actual_row_idx, c_idx) in merged_cells_map:
                    spans = merged_cells_map[(actual_row_idx, c_idx)]
                    if spans["rowspan"] > 1:
                        span_attrs += f' rowspan="{spans["rowspan"]}"'
                    if spans["colspan"] > 1:
                        span_attrs += f' colspan="{spans["colspan"]}"'

                # 生成单元格HTML
                cell_html = self._generate_cell_html(cell, style_class, span_attrs, is_header)
                table_parts.append(cell_html)

            table_parts.append('</tr>')

    def _generate_cell_html(self, cell: Cell, style_class: str, span_attrs: str, is_header: bool = False) -> str:
        """
        生成单个单元格的HTML，支持超链接和注释。

        Args:
            cell: 单元格对象
            style_class: 样式类字符串
            span_attrs: 跨行跨列属性
            is_header: 是否为表头单元格

        Returns:
            单元格HTML字符串
        """
        # 处理单元格值并转义HTML
        formatted_value = self._format_cell_value(cell)
        cell_value = self._escape_html(formatted_value)

        # 处理超链接
        if cell.style and cell.style.hyperlink:
            # 转义超链接URL
            href = self._escape_html(cell.style.hyperlink)
            cell_content = f'<a href="{href}">{cell_value}</a>'
        else:
            cell_content = cell_value

        # 处理注释（工具提示）
        title_attr = ""
        if cell.style and cell.style.comment:
            # 转义注释内容
            comment = self._escape_html(cell.style.comment)
            title_attr = f' title="{comment}"'

        # 处理数字格式（作为data属性）
        data_attr = ""
        if cell.style and cell.style.number_format:
            # 转义数字格式信息
            number_format = self._escape_html(cell.style.number_format)
            data_attr = f' data-number-format="{number_format}"'

        # 选择合适的标签
        tag = 'th' if is_header else 'td'
        return f'<{tag}{style_class}{span_attrs}{title_attr}{data_attr}>{cell_content}</{tag}>'

    def _format_cell_value(self, cell: Cell) -> str:
        """
        格式化单元格值，支持数字和日期格式。

        Args:
            cell: 单元格对象

        Returns:
            格式化后的字符串
        """
        if cell.value is None:
            return ""

        # 如果有数字格式，尝试应用格式化
        if cell.style and cell.style.number_format:
            try:
                return self._apply_number_format(cell.value, cell.style.number_format)
            except Exception:
                # 格式化失败，使用默认格式
                pass

        # 默认格式化
        if isinstance(cell.value, float):
            # 浮点数保留合理的小数位数
            if cell.value.is_integer():
                return str(int(cell.value))
            else:
                return f"{cell.value:.2f}".rstrip('0').rstrip('.')
        elif isinstance(cell.value, int):
            return str(cell.value)
        else:
            return str(cell.value)

    def _apply_number_format(self, value, number_format: str) -> str:
        """
        应用数字格式。

        Args:
            value: 原始值
            number_format: 数字格式字符串

        Returns:
            格式化后的字符串
        """


        # 使用格式映射
        if number_format in NUMBER_FORMAT_MAP:
            return NUMBER_FORMAT_MAP[number_format](value)

        # 检查是否为日期格式
        if isinstance(value, datetime.datetime):
            if "yyyy" in number_format.lower() or "mm" in number_format.lower() or "dd" in number_format.lower():
                format_lower = number_format.lower()
                for excel_fmt, python_fmt in DATE_FORMAT_MAP.items():
                    if excel_fmt in format_lower:
                        return value.strftime(python_fmt)

                # 默认日期格式
                return value.strftime("%Y-%m-%d")



        # 默认返回字符串
        return str(value)

    def _escape_html(self, text: str) -> str:
        """
        转义HTML特殊字符，防止XSS攻击。

        Args:
            text: 原始文本

        Returns:
            转义后的文本
        """
        if not isinstance(text, str):
            text = str(text)

        return (text.replace('&', '&amp;')
                   .replace('<', '&lt;')
                   .replace('>', '&gt;')
                   .replace('"', '&quot;')
                   .replace("'", '&#x27;'))

    def _get_html_template(self) -> str:
        """
        获取HTML模板。
        
        Returns:
            HTML模板字符串
        """
        return """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
{css_styles}
    </style>
</head>
<body>
    <h1>{sheet_name}</h1>
{table_content}
</body>
</html>"""
    
    def _compact_html(self, html: str) -> str:
        """
        压缩HTML内容。
        
        Args:
            html: 原始HTML
            
        Returns:
            压缩后的HTML
        """
        # 简单的压缩：移除多余的空白字符
        lines = html.split('\n')
        compact_lines = [line.strip() for line in lines if line.strip()]
        return '\n'.join(compact_lines)
