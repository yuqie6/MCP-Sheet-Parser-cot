"""
HTML转换器模块

将Sheet对象转换为高保真度的HTML文件，实现95%样式保真度目标。
"""

import logging
from pathlib import Path
from typing import Dict, Any
from src.models.table_model import Sheet, Row, Cell, Style

logger = logging.getLogger(__name__)


class HTMLConverter:
    """HTML转换器，将Sheet对象转换为完美的HTML文件。"""
    
    def __init__(self, compact_mode: bool = False):
        """
        初始化HTML转换器。
        
        Args:
            compact_mode: 是否使用紧凑模式（减少空白字符）
        """
        self.compact_mode = compact_mode
    
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
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        """)
        
        # 生成样式类
        for style_id, style in styles.items():
            css_rule = f".{style_id} {{"
            
            if style.font_name:
                css_rule += f" font-family: {style.font_name};"
            if style.font_size:
                css_rule += f" font-size: {style.font_size}pt;"
            if style.font_color:
                css_rule += f" color: {style.font_color};"
            if style.background_color and style.background_color != "#FFFFFF":
                css_rule += f" background-color: {style.background_color};"
            if style.bold:
                css_rule += " font-weight: bold;"
            if style.italic:
                css_rule += " font-style: italic;"
            if style.underline:
                css_rule += " text-decoration: underline;"
            if style.text_align:
                css_rule += f" text-align: {style.text_align};"
            if style.vertical_align:
                css_rule += f" vertical-align: {style.vertical_align};"
            
            css_rule += " }"
            css_rules.append(css_rule)
        
        return "\n".join(css_rules)
    
    def _generate_table(self, sheet: Sheet, styles: Dict[str, Style]) -> str:
        """
        生成表格HTML。
        
        Args:
            sheet: Sheet对象
            styles: 样式字典
            
        Returns:
            表格HTML字符串
        """
        table_parts = ['<table>']
        
        for row in sheet.rows:
            table_parts.append('<tr>')
            
            for cell in row.cells:
                # 确定样式类
                style_class = ""
                if cell.style:
                    style_key = self._get_style_key(cell.style)
                    for style_id, style_obj in styles.items():
                        if self._get_style_key(style_obj) == style_key:
                            style_class = f' class="{style_id}"'
                            break
                
                # 处理单元格值
                cell_value = str(cell.value) if cell.value is not None else ""
                
                # 生成单元格HTML
                table_parts.append(f'<td{style_class}>{cell_value}</td>')
            
            table_parts.append('</tr>')
        
        table_parts.append('</table>')
        
        return "\n".join(table_parts)
    
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
