import logging
from pathlib import Path
from typing import Any

from src.converters.cell_converter import CellConverter
from src.converters.chart_converter import ChartConverter
from src.converters.style_converter import StyleConverter
from src.converters.table_structure_converter import TableStructureConverter
from src.models.table_model import Sheet
from src.utils.html_utils import compact_html, create_html_element

logger = logging.getLogger(__name__)


class HTMLConverter:
    """将 Sheet 对象转换为 HTML 文件的转换器。"""

    def __init__(self, compact_mode: bool = False, header_rows: int = 1, auto_detect_headers: bool = True):
        """
        初始化 HTML 转换器。
        参数：
            compact_mode: 是否使用紧凑模式（减少空白）。
            header_rows: 表头行数，默认 1。
            auto_detect_headers: 是否自动检测多行表头。
        """
        self.compact_mode = compact_mode
        self.header_rows = header_rows
        self.auto_detect_headers = auto_detect_headers
        self.style_converter = StyleConverter()
        self.cell_converter = CellConverter(self.style_converter)
        self.table_converter = TableStructureConverter(self.cell_converter, self.style_converter)
        self.chart_converter = ChartConverter(self.cell_converter)

    def convert_to_files(self, sheets: list[Sheet], output_path: str) -> list[dict[str, Any]]:
        """
        批量将多个 Sheet 对象转换为多个 HTML 文件。
        参数：
            sheets: Sheet 对象列表。
            output_path: 输出文件路径模板。
        返回：
            转换结果信息列表。
        """
        results = []
        output_p = Path(output_path)
        base_name = output_p.stem
        extension = output_p.suffix or '.html'

        for i, sheet in enumerate(sheets):
            if len(sheets) > 1:
                sheet_output_path = output_p.parent / f"{base_name}-{sheet.name}{extension}"
            else:
                sheet_output_path = output_p
            try:
                html_content = self._generate_html(sheet)
                sheet_output_path.parent.mkdir(parents=True, exist_ok=True)
                with open(sheet_output_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                file_size = sheet_output_path.stat().st_size
                results.append({
                    "status": "success",
                    "output_path": str(sheet_output_path.absolute()),
                    "file_size": file_size,
                    "file_size_kb": round(file_size / 1024, 2),
                    "sheet_name": sheet.name,
                    "rows_converted": len(sheet.rows),
                    "cells_converted": sum(len(row.cells) for row in sheet.rows),
                    "has_styles": any(any(cell.style for cell in row.cells) for row in sheet.rows),
                    "has_merged_cells": len(sheet.merged_cells) > 0
                })
            except Exception as e:
                logger.error(f"HTML conversion failed for sheet '{sheet.name}': {e}")
                results.append({"status": "error", "sheet_name": sheet.name, "error": str(e)})
        return results

    def _generate_html(self, sheet: Sheet) -> str:
        """
        生成完整的 HTML 内容。
        参数：
            sheet: Sheet 对象。
        返回：
            HTML 字符串。
        """
        # 智能检测表头行数
        effective_header_rows = self._detect_header_rows(sheet) if self.auto_detect_headers else self.header_rows

        # 检测是否有Excel标题行
        has_excel_title = self._detect_title_row(sheet)

        styles = self.style_converter.collect_styles(sheet)
        css_content = self.style_converter.generate_css(styles, sheet)
        table_html = self.table_converter.generate_table(sheet, styles, effective_header_rows)
        overlay_charts_html = self.chart_converter.generate_overlay_charts_html(sheet)
        standalone_charts_html = self.chart_converter.generate_standalone_charts_html(sheet.charts)

        # 根据是否检测到Excel标题来决定是否包含<h1>标签
        html_template = self._get_html_template(include_h1=not has_excel_title)
        # 修复：将图表覆盖层放在表格容器内部，确保正确的相对定位
        table_with_charts = create_html_element('div',
                                               f'\n{table_html}\n{overlay_charts_html}\n',
                                               css_classes=['table-container'])
        html_content = html_template.format(
            title=f"Table: {sheet.name}",
            css_styles=css_content,
            table_content=table_with_charts,
            overlay_charts_content="",  # 已经包含在table_content中
            charts_content=standalone_charts_html,
            sheet_name=sheet.name
        )
        if self.compact_mode:
            html_content = compact_html(html_content)
        return html_content

    def _detect_header_rows(self, sheet: Sheet) -> int:
        """
        智能检测表头行数（样式、内容、合并单元格等多种启发式规则）。
        参数：sheet: Sheet对象
        返回：检测到的表头行数
        """
        if not sheet.rows or len(sheet.rows) < 2:
            return min(1, len(sheet.rows))  # 如果只有0-1行，返回实际行数

        # 默认至少有1行表头
        header_rows = 1
        max_check_rows = min(10, len(sheet.rows))  # 最多检查前10行

        # 检查合并单元格
        header_merged_cells = []
        for merged_range in sheet.merged_cells:
            try:
                from src.utils.range_parser import parse_range_string
                start_row, _, end_row, _ = parse_range_string(merged_range)
                if start_row < max_check_rows:
                    header_merged_cells.append((start_row, end_row))
            except Exception:
                pass

        # 如果有跨行的合并单元格，可能是多行表头
        for start_row, end_row in header_merged_cells:
            if end_row >= header_rows:
                header_rows = end_row + 1

        # 检查样式差异
        data_row_styles = []
        header_row_styles = []

        # 收集前几行的样式特征
        for i in range(min(max_check_rows, len(sheet.rows))):
            row = sheet.rows[i]
            style_features = self._get_row_style_features(row)
            if i < header_rows:
                header_row_styles.append(style_features)
            else:
                data_row_styles.append(style_features)

        # 如果数据行样式与表头行显著不同，可能是多行表头
        if data_row_styles and header_row_styles:
            for i in range(header_rows, min(max_check_rows, len(sheet.rows))):
                if self._is_header_style(sheet.rows[i], header_row_styles, data_row_styles):
                    header_rows = i + 1
                else:
                    break

        return header_rows

    def _detect_title_row(self, sheet: Sheet) -> bool:
        """
        检测第一行是否为Excel标题行。

        检测标准：
        1. 强标题：字体大小 >= 18pt（无论其他样式）
        2. 中等标题：字体大小 >= 14pt 且 (粗体 且 居中对齐)
        3. 弱标题：字体大小 >= 16pt 且 (粗体 或 居中对齐)

        参数：
            sheet: Sheet 对象
        返回：
            bool: 如果检测到标题行返回True，否则返回False
        """
        if not sheet.rows or not sheet.rows[0].cells:
            return False

        first_row = sheet.rows[0]

        # 检查第一行的所有单元格
        for cell in first_row.cells:
            if cell.value and str(cell.value).strip() and cell.style:
                font_size = cell.style.font_size or 0
                is_bold = cell.style.bold or False
                is_center = cell.style.text_align == 'center'

                # 强标题：字体大小 >= 18pt
                if font_size >= 18:
                    return True

                # 中等标题：字体大小 >= 14pt 且 粗体 且 居中
                if font_size >= 14 and is_bold and is_center:
                    return True

                # 弱标题：字体大小 >= 16pt 且 (粗体 或 居中)
                if font_size >= 16 and (is_bold or is_center):
                    return True

        return False

    def _get_row_style_features(self, row) -> dict:
        """提取行的样式特征。"""
        features = {
            'bold_count': 0,
            'bg_color_count': 0,
            'text_count': 0,
            'number_count': 0,
            'empty_count': 0
        }

        for cell in row.cells:
            # 检查样式
            if cell.style:
                if cell.style.bold:
                    features['bold_count'] += 1
                if cell.style.background_color:
                    features['bg_color_count'] += 1

            # 检查内容类型
            if cell.value is None or cell.value == '':
                features['empty_count'] += 1
            elif isinstance(cell.value, (int, float)):
                features['number_count'] += 1
            else:
                features['text_count'] += 1

        return features

    def _is_header_style(self, row, header_styles, data_styles) -> bool:
        """判断行是否更像表头样式。"""
        row_features = self._get_row_style_features(row)

        # 计算与表头样式和数据样式的相似度
        header_similarity = 0
        data_similarity = 0

        for header_style in header_styles:
            similarity = 0
            for key in header_style:
                if header_style[key] > 0 and row_features[key] > 0:
                    similarity += 1
            header_similarity = max(header_similarity, similarity)

        for data_style in data_styles:
            similarity = 0
            for key in data_style:
                if data_style[key] > 0 and row_features[key] > 0:
                    similarity += 1
            data_similarity = max(data_similarity, similarity)

        # 如果行更像表头，或者包含大量粗体/背景色，认为是表头
        return (header_similarity > data_similarity or
                row_features['bold_count'] > len(row.cells) * 0.5 or
                row_features['bg_color_count'] > len(row.cells) * 0.5)

    def _get_html_template(self, include_h1: bool = True) -> str:
        """
        获取 HTML 模板字符串。

        参数：
            include_h1: 是否包含<h1>标签，默认为True
        """
        # 根据参数决定是否包含<h1>标签
        h1_section = "    <h1>{sheet_name}</h1>\n" if include_h1 else ""

        template = """<!DOCTYPE html>
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
""" + h1_section + """{table_content}
{overlay_charts_content}
{charts_content}
    <script>
        // 强制应用文字溢出样式
        document.addEventListener('DOMContentLoaded', function() {{
            const overflowCells = document.querySelectorAll('.text-overflow');
            overflowCells.forEach(function(cell) {{
                cell.style.setProperty('overflow', 'visible', 'important');
                cell.style.setProperty('white-space', 'nowrap', 'important');
                cell.style.setProperty('width', 'auto', 'important');
                cell.style.setProperty('min-width', 'auto', 'important');
                cell.style.setProperty('word-wrap', 'normal', 'important');
                cell.style.setProperty('position', 'relative', 'important');
                cell.style.setProperty('z-index', '5', 'important');
                console.log('Applied overflow styles to:', cell.textContent.trim());
            }});
        }});
    </script>
</body>
</html>"""

        return template
