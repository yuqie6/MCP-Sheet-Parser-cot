"""
性能优化工具模块

提供文件大小估算、性能监控和智能处理建议功能。
"""

import time
from typing import Any
from dataclasses import dataclass
from src.models.table_model import Sheet


@dataclass
class PerformanceMetrics:
    """性能指标数据类。"""
    rows: int
    cols: int
    total_cells: int
    non_empty_cells: int
    styled_cells: int
    estimated_html_size: int
    estimated_json_size: int
    processing_time: float
    recommendation: str


class PerformanceOptimizer:
    """性能优化器，提供文件大小估算和处理建议。"""
    
    # 性能阈值配置
    SMALL_FILE_THRESHOLD = 1000      # 小文件：< 1000个单元格
    MEDIUM_FILE_THRESHOLD = 10000    # 中文件：< 10000个单元格
    LARGE_FILE_THRESHOLD = 50000     # 大文件：< 50000个单元格
    
    HTML_SIZE_THRESHOLD_DIRECT = 50000    # 直接返回HTML的大小阈值
    HTML_SIZE_THRESHOLD_WARNING = 100000  # 警告阈值
    
    def __init__(self):
        """初始化性能优化器。"""
        self.start_time = None
    
    def start_timing(self) -> None:
        """开始计时。"""
        self.start_time = time.time()
    
    def end_timing(self) -> float:
        """结束计时并返回耗时。"""
        if self.start_time is None:
            return 0.0
        return time.time() - self.start_time
    
    def analyze_sheet_performance(self, sheet: Sheet) -> PerformanceMetrics:
        """
        分析表格的性能指标。
        
        Args:
            sheet: 要分析的表格对象
            
        Returns:
            性能指标对象
        """
        # 基本统计
        rows = len(sheet.rows)
        cols = len(sheet.rows[0].cells) if sheet.rows else 0
        total_cells = rows * cols
        
        # 统计非空和有样式的单元格
        non_empty_cells = 0
        styled_cells = 0
        
        for row in sheet.rows:
            for cell in row.cells:
                if cell.value is not None and str(cell.value).strip():
                    non_empty_cells += 1
                if cell.style and self._has_custom_style(cell.style):
                    styled_cells += 1
        
        # 估算HTML大小
        estimated_html_size = self._estimate_html_size(sheet, non_empty_cells, styled_cells)
        
        # 估算JSON大小
        estimated_json_size = self._estimate_json_size(sheet, non_empty_cells, styled_cells)
        
        # 获取处理建议
        recommendation = self._get_processing_recommendation(
            total_cells, estimated_html_size, non_empty_cells
        )
        
        # 获取处理时间
        processing_time = self.end_timing()
        
        return PerformanceMetrics(
            rows=rows,
            cols=cols,
            total_cells=total_cells,
            non_empty_cells=non_empty_cells,
            styled_cells=styled_cells,
            estimated_html_size=estimated_html_size,
            estimated_json_size=estimated_json_size,
            processing_time=processing_time,
            recommendation=recommendation
        )
    
    def _has_custom_style(self, style) -> bool:
        """检查是否有自定义样式。"""
        if not style:
            return False

        # 检查是否有非默认样式
        return (
            style.bold or style.italic or style.underline or
            style.font_color != "#000000" or
            style.background_color != "#FFFFFF" or
            style.text_align != "left" or
            style.border_top or style.border_bottom or
            style.border_left or style.border_right or
            style.hyperlink is not None or style.comment is not None
        )
    
    def _estimate_html_size(self, sheet: Sheet, non_empty_cells: int, styled_cells: int) -> int:
        """
        估算HTML输出大小。
        
        Args:
            sheet: 表格对象
            non_empty_cells: 非空单元格数量
            styled_cells: 有样式的单元格数量
            
        Returns:
            估算的HTML大小（字符数）
        """
        # 基础HTML结构开销
        base_html_overhead = 500  # HTML头部、表格标签等
        
        # 每个单元格的基础开销
        cell_base_overhead = 20  # <td></td> 标签
        
        # 样式开销
        style_overhead_per_cell = 50  # 内联样式或CSS类
        
        # 内容大小估算
        content_size = 0
        for row in sheet.rows:
            for cell in row.cells:
                if cell.value is not None:
                    content_size += len(str(cell.value))
        
        # 总大小估算
        total_size = (
            base_html_overhead +
            len(sheet.rows) * len(sheet.rows[0].cells if sheet.rows else []) * cell_base_overhead +
            styled_cells * style_overhead_per_cell +
            content_size
        )
        
        return total_size
    
    def _estimate_json_size(self, sheet: Sheet, non_empty_cells: int, styled_cells: int) -> int:
        """
        估算JSON输出大小。
        
        Args:
            sheet: 表格对象
            non_empty_cells: 非空单元格数量
            styled_cells: 有样式的单元格数量
            
        Returns:
            估算的JSON大小（字符数）
        """
        # JSON结构开销
        base_json_overhead = 100  # 元数据、数组结构等

        # 每个单元格的JSON开销
        cell_json_overhead = 15  # JSON对象结构

        # 样式JSON开销
        style_json_overhead = 50  # 样式对象
        
        # 内容大小
        content_size = 0
        for row in sheet.rows:
            for cell in row.cells:
                if cell.value is not None:
                    content_size += len(str(cell.value)) + 5  # JSON字符串引号等
        
        # 总大小估算
        total_size = (
            base_json_overhead +
            non_empty_cells * cell_json_overhead +
            styled_cells * style_json_overhead +
            content_size
        )
        
        return total_size
    
    def _get_processing_recommendation(self, total_cells: int, estimated_html_size: int, 
                                     non_empty_cells: int) -> str:
        """
        获取处理建议。
        
        Args:
            total_cells: 总单元格数
            estimated_html_size: 估算HTML大小
            non_empty_cells: 非空单元格数
            
        Returns:
            处理建议字符串
        """
        if total_cells < self.SMALL_FILE_THRESHOLD:
            return "small_file_direct"
        elif total_cells < self.MEDIUM_FILE_THRESHOLD:
            if estimated_html_size < self.HTML_SIZE_THRESHOLD_DIRECT:
                return "medium_file_direct"
            else:
                return "medium_file_with_warning"
        elif total_cells < self.LARGE_FILE_THRESHOLD:
            return "large_file_recommend_file_output"
        else:
            return "very_large_file_require_pagination"
    
    def get_recommendation_message(self, recommendation: str, metrics: PerformanceMetrics) -> str:
        """
        获取用户友好的建议信息。
        
        Args:
            recommendation: 建议类型
            metrics: 性能指标
            
        Returns:
            用户友好的建议信息
        """
        messages = {
            "small_file_direct": f"小文件（{metrics.total_cells:,}个单元格）- 建议直接返回HTML内容",
            "medium_file_direct": f"中等文件（{metrics.total_cells:,}个单元格）- 可以直接返回HTML内容",
            "medium_file_with_warning": f"中等文件（{metrics.total_cells:,}个单元格）- 建议使用文件输出模式以获得更好性能",
            "large_file_recommend_file_output": f"大文件（{metrics.total_cells:,}个单元格）- 强烈建议使用文件输出模式",
            "very_large_file_require_pagination": f"超大文件（{metrics.total_cells:,}个单元格）- 需要使用分页处理"
        }
        
        return messages.get(recommendation, "未知文件大小 - 建议谨慎处理")
    
    def should_use_pagination(self, metrics: PerformanceMetrics) -> bool:
        """
        判断是否应该使用分页。
        
        Args:
            metrics: 性能指标
            
        Returns:
            是否应该使用分页
        """
        return metrics.total_cells >= self.LARGE_FILE_THRESHOLD
    
    def calculate_pagination_params(self, sheet: Sheet, max_rows_per_page: int = 1000) -> dict[str, Any]:
        """
        计算分页参数。
        
        Args:
            sheet: 表格对象
            max_rows_per_page: 每页最大行数
            
        Returns:
            分页参数字典
        """
        total_rows = len(sheet.rows)
        total_pages = (total_rows + max_rows_per_page - 1) // max_rows_per_page
        
        return {
            "total_rows": total_rows,
            "max_rows_per_page": max_rows_per_page,
            "total_pages": total_pages,
            "pages": [
                {
                    "page_number": i + 1,
                    "start_row": i * max_rows_per_page,
                    "end_row": min((i + 1) * max_rows_per_page, total_rows),
                    "row_count": min(max_rows_per_page, total_rows - i * max_rows_per_page)
                }
                for i in range(total_pages)
            ]
        }
