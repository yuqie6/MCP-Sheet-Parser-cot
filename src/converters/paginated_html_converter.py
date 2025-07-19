"""
分页HTML转换器

支持大文件的分页处理和性能优化
"""

import logging
from typing import Any

from .html_converter import HTMLConverter
from ..models.table_model import Sheet

logger = logging.getLogger(__name__)


class PaginatedHTMLConverter(HTMLConverter):
    """分页HTML转换器，继承自HTMLConverter，添加分页功能。"""
    
    def __init__(self, compact_mode: bool = False, page_size: int = 100, page_number: int = 1, header_rows: int = 1, auto_detect_headers: bool = True):
        """
        初始化分页HTML转换器。

        参数：
            compact_mode: 是否使用紧凑模式
            page_size: 每页显示的行数
            page_number: 当前页码（从1开始）
            header_rows: 表头行数，默认第一行为表头
            auto_detect_headers: 是否自动检测多行表头
        """
        super().__init__(compact_mode, header_rows, auto_detect_headers)
        self.page_size = max(1, page_size)  # 确保页面大小至少为1
        self.page_number = max(1, page_number)  # 确保页码至少为1
        
    def _generate_html(self, sheet: Sheet) -> str:
        """
        生成分页HTML内容。

        参数：
            sheet: Sheet对象

        返回：
            HTML字符串
        """
        # 计算分页信息
        total_rows = len(sheet.rows)
        total_pages = (total_rows + self.page_size - 1) // self.page_size if total_rows > 0 else 1
        
        # 确保页码在有效范围内
        current_page = min(self.page_number, total_pages)
        
        # 计算当前页的数据范围
        start_row = (current_page - 1) * self.page_size
        end_row = min(start_row + self.page_size, total_rows)
        
        # 创建分页后的Sheet对象
        paginated_sheet = self._create_paginated_sheet(sheet, start_row, end_row)
        
        # 生成基础HTML
        base_html = super()._generate_html(paginated_sheet)
        
        # 添加分页导航
        pagination_html = self._generate_pagination_controls(
            current_page, total_pages, total_rows, start_row, end_row
        )
        
        # 插入分页控件到HTML中
        return self._insert_pagination_controls(base_html, pagination_html, sheet.name)
    
    def _create_paginated_sheet(self, sheet: Sheet, start_row: int, end_row: int) -> Sheet:
        """
        创建分页后的Sheet对象。

        参数：
            sheet: 原始Sheet对象
            start_row: 起始行索引
            end_row: 结束行索引

        返回：
            分页后的Sheet对象
        """
        # 获取当前页的行数据
        paginated_rows = sheet.rows[start_row:end_row] if sheet.rows else []
        
        # 创建新的Sheet对象
        paginated_sheet = Sheet(
            name=sheet.name,
            rows=paginated_rows,
            merged_cells=sheet.merged_cells,  # 保留合并单元格信息
            charts=sheet.charts,  # 保留图表信息
            column_widths=sheet.column_widths,  # 保留列宽信息
            row_heights=sheet.row_heights,  # 保留行高信息
            default_column_width=sheet.default_column_width,
            default_row_height=sheet.default_row_height
        )
        
        return paginated_sheet

    def convert_to_file(self, sheet: Sheet, output_path: str) -> dict[str, Any]:
        """
        将单个Sheet对象转换为分页HTML文件。

        参数：
            sheet: Sheet对象
            output_path: 输出文件路径

        返回：
            转换结果信息字典
        """
        from pathlib import Path

        try:
            # 生成HTML内容
            html_content = self._generate_html(sheet)

            # 写入文件
            output_p = Path(output_path)
            output_p.parent.mkdir(parents=True, exist_ok=True)

            with open(output_p, 'w', encoding='utf-8') as f:
                f.write(html_content)

            # 计算文件大小
            file_size = output_p.stat().st_size

            return {
                "status": "success",
                "output_path": str(output_p.absolute()),
                "file_size": file_size,
                "file_size_kb": round(file_size / 1024, 2),
                "sheet_name": sheet.name,
                "rows_converted": len(sheet.rows),
                "cells_converted": sum(len(row.cells) for row in sheet.rows),
                "has_styles": any(any(cell.style for cell in row.cells) for row in sheet.rows),
                "has_merged_cells": len(sheet.merged_cells) > 0,
                "page_size": self.page_size,
                "page_number": self.page_number,
                "total_pages": (len(sheet.rows) + self.page_size - 1) // self.page_size if len(sheet.rows) > 0 else 1
            }

        except Exception as e:
            return {
                "status": "error",
                "error": str(e),
                "sheet_name": sheet.name
            }

    def _generate_pagination_controls(self, current_page: int, total_pages: int,
                                    total_rows: int, start_row: int, end_row: int) -> str:
        """
        生成分页导航控件HTML。

        参数：
            current_page: 当前页码
            total_pages: 总页数
            total_rows: 总行数
            start_row: 当前页起始行
            end_row: 当前页结束行

        返回：
            分页控件HTML字符串
        """
        pagination_parts = []
        
        # 分页信息
        pagination_parts.append('<div class="pagination-info">')
        pagination_parts.append(f'<span>显示第 {start_row + 1}-{end_row} 行，共 {total_rows} 行</span>')
        pagination_parts.append(f'<span>第 {current_page} 页，共 {total_pages} 页</span>')
        pagination_parts.append('</div>')
        
        # 分页按钮
        pagination_parts.append('<div class="pagination-controls">')
        
        # 上一页按钮
        if current_page > 1:
            pagination_parts.append(f'<button onclick="loadPage({current_page - 1})" class="page-btn">上一页</button>')
        else:
            pagination_parts.append('<button class="page-btn disabled" disabled>上一页</button>')
        
        # 页码按钮（显示当前页附近的页码）
        start_page = max(1, current_page - 2)
        end_page = min(total_pages, current_page + 2)
        
        if start_page > 1:
            pagination_parts.append('<button onclick="loadPage(1)" class="page-btn">1</button>')
            if start_page > 2:
                pagination_parts.append('<span class="page-ellipsis">...</span>')
        
        for page in range(start_page, end_page + 1):
            if page == current_page:
                pagination_parts.append(f'<button class="page-btn current">{page}</button>')
            else:
                pagination_parts.append(f'<button onclick="loadPage({page})" class="page-btn">{page}</button>')
        
        if end_page < total_pages:
            if end_page < total_pages - 1:
                pagination_parts.append('<span class="page-ellipsis">...</span>')
            pagination_parts.append(f'<button onclick="loadPage({total_pages})" class="page-btn">{total_pages}</button>')
        
        # 下一页按钮
        if current_page < total_pages:
            pagination_parts.append(f'<button onclick="loadPage({current_page + 1})" class="page-btn">下一页</button>')
        else:
            pagination_parts.append('<button class="page-btn disabled" disabled>下一页</button>')
        
        pagination_parts.append('</div>')
        
        return '\n'.join(pagination_parts)
    
    def _insert_pagination_controls(self, base_html: str, pagination_html: str, sheet_name: str) -> str:
        """
        将分页控件插入到HTML中。

        参数：
            base_html: 基础HTML内容
            pagination_html: 分页控件HTML
            sheet_name: 表格名称

        返回：
            包含分页控件的完整HTML
        """
        # 添加分页样式
        pagination_css = self._generate_pagination_css()
        
        # 添加分页JavaScript
        pagination_js = self._generate_pagination_js(sheet_name)
        
        # 在</style>标签前插入分页CSS
        if '</style>' in base_html:
            base_html = base_html.replace('</style>', f'{pagination_css}\n</style>')
        
        # 在<h1>标签后插入顶部分页控件
        if f'<h1>{sheet_name}</h1>' in base_html:
            base_html = base_html.replace(
                f'<h1>{sheet_name}</h1>',
                f'<h1>{sheet_name}</h1>\n<div class="pagination-container top">\n{pagination_html}\n</div>'
            )
        
        # 在</table>标签后插入底部分页控件
        if '</table>' in base_html:
            base_html = base_html.replace(
                '</table>',
                f'</table>\n<div class="pagination-container bottom">\n{pagination_html}\n</div>'
            )
        
        # 在</body>标签前插入JavaScript
        if '</body>' in base_html:
            base_html = base_html.replace('</body>', f'{pagination_js}\n</body>')
        
        return base_html
    
    def _generate_pagination_css(self) -> str:
        """生成分页控件的CSS样式。"""
        return """
        /* 分页控件样式 */
        .pagination-container {
            margin: 20px 0;
            text-align: center;
            border: 1px solid #ddd;
            border-radius: 5px;
            padding: 15px;
            background-color: #f9f9f9;
        }
        
        .pagination-info {
            margin-bottom: 10px;
            color: #666;
            font-size: 14px;
        }
        
        .pagination-info span {
            margin: 0 10px;
        }
        
        .pagination-controls {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 5px;
            flex-wrap: wrap;
        }
        
        .page-btn {
            padding: 8px 12px;
            border: 1px solid #ddd;
            background-color: white;
            color: #333;
            cursor: pointer;
            border-radius: 3px;
            font-size: 14px;
            transition: all 0.2s;
        }
        
        .page-btn:hover:not(.disabled) {
            background-color: #e9ecef;
            border-color: #adb5bd;
        }
        
        .page-btn.current {
            background-color: #007bff;
            color: white;
            border-color: #007bff;
        }
        
        .page-btn.disabled {
            background-color: #f8f9fa;
            color: #6c757d;
            cursor: not-allowed;
            opacity: 0.6;
        }
        
        .page-ellipsis {
            padding: 8px 4px;
            color: #6c757d;
        }
        
        @media (max-width: 768px) {
            .pagination-controls {
                gap: 3px;
            }
            
            .page-btn {
                padding: 6px 8px;
                font-size: 12px;
            }
        }
        """
    
    def _generate_pagination_js(self, sheet_name: str) -> str:
        """生成分页控件的JavaScript代码。"""
        return f"""
        <script>
        function loadPage(pageNumber) {{
            // 这里应该调用MCP工具重新生成指定页面的HTML
            // 由于这是静态HTML，我们只能显示提示信息
            alert('要加载第 ' + pageNumber + ' 页，请使用MCP工具调用：\\n' +
                  'convert_to_html(file_path="{sheet_name}", page_number=' + pageNumber + ')');
        }}
        
        // 键盘导航支持
        document.addEventListener('keydown', function(e) {{
            if (e.key === 'ArrowLeft') {{
                const prevBtn = document.querySelector('.page-btn:not(.disabled)[onclick*="' + (getCurrentPage() - 1) + '"]');
                if (prevBtn) prevBtn.click();
            }} else if (e.key === 'ArrowRight') {{
                const nextBtn = document.querySelector('.page-btn:not(.disabled)[onclick*="' + (getCurrentPage() + 1) + '"]');
                if (nextBtn) nextBtn.click();
            }}
        }});
        
        function getCurrentPage() {{
            const currentBtn = document.querySelector('.page-btn.current');
            return currentBtn ? parseInt(currentBtn.textContent) : 1;
        }}
        </script>
        """
