import logging
from pathlib import Path
from typing import List, Any

from src.converters.cell_converter import CellConverter
from src.converters.chart_converter import ChartConverter
from src.converters.style_converter import StyleConverter
from src.converters.table_structure_converter import TableStructureConverter
from src.models.table_model import Sheet
from src.utils.html_utils import compact_html

logger = logging.getLogger(__name__)


class HTMLConverter:
    """A converter that transforms Sheet objects into HTML files."""

    def __init__(self, compact_mode: bool = False, header_rows: int = 1):
        """
        Initializes the HTML converter.
        Args:
            compact_mode: Whether to use compact mode (reduces whitespace).
            header_rows: The number of header rows, default is 1.
        """
        self.compact_mode = compact_mode
        self.header_rows = header_rows
        self.style_converter = StyleConverter()
        self.cell_converter = CellConverter(self.style_converter)
        self.table_converter = TableStructureConverter(self.cell_converter, self.style_converter)
        self.chart_converter = ChartConverter(self.cell_converter)

    def convert_to_files(self, sheets: List[Sheet], output_path: str) -> List[dict[str, Any]]:
        """
        Converts multiple Sheet objects to multiple HTML files.
        Args:
            sheets: A list of Sheet objects.
            output_path: The output file path template.
        Returns:
            A list of conversion result information.
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
        Generates the complete HTML content.
        Args:
            sheet: The Sheet object.
        Returns:
            The HTML string.
        """
        styles = self.style_converter.collect_styles(sheet)
        css_content = self.style_converter.generate_css(styles, sheet)
        table_html = self.table_converter.generate_table(sheet, styles, self.header_rows)
        overlay_charts_html = self.chart_converter.generate_overlay_charts_html(sheet)
        standalone_charts_html = self.chart_converter.generate_standalone_charts_html(sheet.charts)

        html_template = self._get_html_template()
        # 修复：将图表覆盖层放在表格容器内部，确保正确的相对定位
        table_with_charts = f'<div class="table-container">\n{table_html}\n{overlay_charts_html}\n</div>'
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

    def _get_html_template(self) -> str:
        """
        Gets the HTML template.
        Returns:
            The HTML template string.
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
