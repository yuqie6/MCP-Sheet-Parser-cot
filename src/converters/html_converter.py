"""
HTML转换器模块

将Sheet对象转换为高保真度的HTML文件，实现95%样式保真度目标。
"""

import logging
from pathlib import Path
from typing import Dict, Any, Set, Tuple
from src.models.table_model import Sheet, Cell, Style
from src.utils.range_parser import parse_range_string

logger = logging.getLogger(__name__)



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
    
    def convert_to_file(self, sheet: Sheet, output_path: str) -> Dict[str, Any]:
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
    
    def _collect_styles(self, sheet: Sheet) -> Dict[str, Style]:
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
    
    def _generate_css(self, styles: Dict[str, Style]) -> str:
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
                css_rule += f" font-family: {style.font_name};"
            if style.font_size:
                css_rule += f" font-size: {style.font_size}pt;"
            if style.font_color:
                css_rule += f" color: {style.font_color} !important;"
            if style.bold:
                css_rule += " font-weight: bold;"
            if style.italic:
                css_rule += " font-style: italic;"
            if style.underline:
                css_rule += " text-decoration: underline;"

            # 背景和填充
            if style.background_color:
                css_rule += f" background-color: {style.background_color};"

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
            border_width = self._parse_border_style(style.border_top)
            border_css += f" border-top: {border_width} solid {border_color};"

        if style.border_bottom:
            border_width = self._parse_border_style(style.border_bottom)
            border_css += f" border-bottom: {border_width} solid {border_color};"

        if style.border_left:
            border_width = self._parse_border_style(style.border_left)
            border_css += f" border-left: {border_width} solid {border_color};"

        if style.border_right:
            border_width = self._parse_border_style(style.border_right)
            border_css += f" border-right: {border_width} solid {border_color};"

        return border_css

    def _parse_border_style(self, border_style: str) -> str:
        """
        解析边框样式字符串，转换为CSS边框宽度。

        Args:
            border_style: 边框样式字符串

        Returns:
            CSS边框宽度
        """
        if not border_style:
            return "1px"

        # 简单的边框样式映射
        style_map = {
            "thin": "1px",
            "medium": "2px",
            "thick": "3px",
            "solid": "1px",
            "dashed": "1px",
            "dotted": "1px"
        }

        # 如果是已知样式，返回对应宽度
        if border_style.lower() in style_map:
            return style_map[border_style.lower()]

        # 如果包含数字，尝试提取
        import re
        match = re.search(r'(\d+)', border_style)
        if match:
            return f"{match.group(1)}px"

        # 默认返回1px
        return "1px"

    def _generate_table(self, sheet: Sheet, styles: Dict[str, Style]) -> str:
        """
        生成表格HTML。
        
        Args:
            sheet: Sheet对象
            styles: 样式字典
            
        Returns:
            表格HTML字符串
        """
        
        # 处理合并单元格
        occupied_cells: Set[Tuple[int, int]] = set()
        merged_cells_map: Dict[Tuple[int, int], Dict[str, int]] = {}

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

        table_parts = ['<table>']
        
        # Pre-compute a reverse map from style_key to style_id for efficiency
        style_key_to_id_map = {self._get_style_key(style_obj): style_id for style_id, style_obj in styles.items()}

        for r_idx, row in enumerate(sheet.rows):
            table_parts.append('<tr>')
            
            for c_idx, cell in enumerate(row.cells):
                if (r_idx, c_idx) in occupied_cells:
                    continue

                # 确定样式类
                style_class = ""
                if cell.style:
                    style_key = self._get_style_key(cell.style)
                    style_id = style_key_to_id_map.get(style_key)
                    if style_id:
                        style_class = f' class="{style_id}"'
                
                span_attrs = ""
                if (r_idx, c_idx) in merged_cells_map:
                    spans = merged_cells_map[(r_idx, c_idx)]
                    if spans["rowspan"] > 1:
                        span_attrs += f' rowspan="{spans["rowspan"]}"'
                    if spans["colspan"] > 1:
                        span_attrs += f' colspan="{spans["colspan"]}"'

                # 生成单元格HTML（支持超链接和注释）
                cell_html = self._generate_cell_html(cell, style_class, span_attrs)
                table_parts.append(cell_html)
            
            table_parts.append('</tr>')
        
        table_parts.append('</table>')
        
        return "\n".join(table_parts)

    def _generate_cell_html(self, cell: Cell, style_class: str, span_attrs: str) -> str:
        """
        生成单个单元格的HTML，支持超链接和注释。

        Args:
            cell: 单元格对象
            style_class: 样式类字符串

        Returns:
            单元格HTML字符串
        """
        # 处理单元格值并转义HTML
        cell_value = self._escape_html(str(cell.value) if cell.value is not None else "")

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

        return f'<td{style_class}{span_attrs}{title_attr}{data_attr}>{cell_content}</td>'

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
