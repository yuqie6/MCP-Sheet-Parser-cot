"""
图表定位计算器模块

将Excel图表的EMU定位数据转换为CSS定位信息，实现精确的图表覆盖定位。
"""

from typing import Tuple, Optional
from dataclasses import dataclass
from src.models.table_model import Sheet, ChartPosition


@dataclass
class ChartCSSPosition:
    """图表的CSS定位信息。"""
    left: float  # 从表格左边缘的像素距离
    top: float   # 从表格顶部的点距离
    width: float # 图表宽度（像素）
    height: float # 图表高度（点）


class ChartPositionCalculator:
    """
    图表定位计算器，负责将Excel的EMU坐标系转换为HTML/CSS定位。
    
    Excel使用以下坐标系统：
    - 列宽：Excel单位（字符宽度）
    - 行高：点（pt）
    - 单元格内偏移：EMU（English Metric Units）
    - 1英寸 = 914,400 EMUs = 96像素（96 DPI）
    """
    
    # 常量定义
    EXCEL_TO_PX = 8.43  # Excel列宽单位到像素的转换系数
    EMU_TO_PX = 96 / 914400  # EMU到像素的转换系数（96 DPI）
    
    def __init__(self, sheet: Sheet):
        """
        初始化计算器。
        
        参数：
            sheet: 工作表对象，包含列宽和行高信息
        """
        self.sheet = sheet
        
    def calculate_chart_css_position(self, position: ChartPosition) -> ChartCSSPosition:
        """
        计算图表的CSS定位信息。
        
        参数：
            position: 图表的Excel定位信息
            
        返回：
            ChartCSSPosition: CSS定位信息
        """
        # 使用更保守的定位策略，避免图表超出可视区域
        # 计算相对于表格的位置，而不是绝对累计位置
        
        # 计算起始列的相对位置（大幅减少偏移）
        start_col_width = self.sheet.column_widths.get(position.from_col, self.sheet.default_column_width)
        relative_x = position.from_col * start_col_width * self.EXCEL_TO_PX * 0.1  # 大幅减少为10%
        
        # 计算起始行的相对位置
        start_row_height = self.sheet.row_heights.get(position.from_row, self.sheet.default_row_height)
        relative_y = position.from_row * start_row_height * 0.3  # 减少垂直偏移
        
        # 添加单元格内偏移（转换并大幅限制）
        col_offset_px = position.from_col_offset * self.EMU_TO_PX * 0.1  # 大幅减少EMU偏移影响
        row_offset_px = position.from_row_offset * self.EMU_TO_PX * 0.1  # 大幅减少EMU偏移影响
        
        # 限制偏移量，防止图表跑出可视区域
        max_offset_x = 200  # 减少最大水平偏移
        max_offset_y = 100  # 减少最大垂直偏移
        
        final_x = min(relative_x + col_offset_px, max_offset_x)
        final_y = min(relative_y + row_offset_px, max_offset_y)
        
        # 计算图表尺寸（使用Excel原始尺寸信息）
        # 优先使用EMU尺寸，如果不可用则回退到行列跨度计算
        width = self._calculate_chart_width(position)
        height = self._calculate_chart_height(position)
        
        return ChartCSSPosition(
            left=final_x,
            top=final_y,
            width=width,
            height=height
        )
    
    def _calculate_cell_position(self, col: int, row: int) -> Tuple[float, float]:
        """
        计算指定单元格的累计位置。
        
        参数：
            col: 列索引（0基）
            row: 行索引（0基）
            
        返回：
            Tuple[float, float]: (x_像素, y_点) 位置
        """
        # 计算X位置（累计列宽）
        cumulative_width = 0
        for col_idx in range(col):
            width = self.sheet.column_widths.get(col_idx, self.sheet.default_column_width)
            cumulative_width += width * self.EXCEL_TO_PX
            
        # 计算Y位置（累计行高）
        cumulative_height = 0
        for row_idx in range(row):
            height = self.sheet.row_heights.get(row_idx, self.sheet.default_row_height)
            cumulative_height += height  # 已经是点单位
            
        return cumulative_width, cumulative_height
    
    def _calculate_chart_width(self, position: ChartPosition) -> float:
        """
        计算图表宽度。
        
        参数：
            position: 图表定位信息
            
        返回：
            float: 图表宽度（像素）
        """
        # 计算跨越的列数和总宽度
        col_span = max(1, position.to_col - position.from_col + 1)
        
        # 使用平均列宽计算，避免过大的宽度
        avg_col_width = self.sheet.default_column_width
        if self.sheet.column_widths:
            total_width = sum(self.sheet.column_widths.values())
            avg_col_width = total_width / len(self.sheet.column_widths)
        
        # 基础宽度 + 列偏移
        base_width = col_span * avg_col_width * self.EXCEL_TO_PX * 0.7  # 减少宽度倍数
        col_offset_width = (position.to_col_offset - position.from_col_offset) * self.EMU_TO_PX
        
        # 限制最大宽度，避免图表过大
        max_width = 500
        total_width = min(base_width + col_offset_width, max_width)
        
        return max(200, total_width)  # 最小宽度200px
    
    def _calculate_chart_height(self, position: ChartPosition) -> float:
        """
        计算图表高度。
        
        参数：
            position: 图表定位信息
            
        返回：
            float: 图表高度（点）
        """
        # 计算跨越的行数和总高度
        row_span = max(1, position.to_row - position.from_row + 1)
        
        # 使用平均行高计算
        avg_row_height = self.sheet.default_row_height
        if self.sheet.row_heights:
            total_height = sum(self.sheet.row_heights.values())
            avg_row_height = total_height / len(self.sheet.row_heights)
        
        # 基础高度 + 行偏移（转换为点）
        base_height = row_span * avg_row_height * 0.8  # 减少高度倍数
        row_offset_height = (position.to_row_offset - position.from_row_offset) * self.EMU_TO_PX * 0.75
        
        # 限制最大高度
        max_height = 300
        total_height = min(base_height + row_offset_height, max_height)
        
        return max(150, total_height)  # 最小高度150pt
    
    def get_chart_overlay_css(self, position: ChartPosition, container_id: str = "table-container") -> str:
        """
        生成图表覆盖的CSS样式。
        
        参数：
            position: 图表定位信息
            container_id: 表格容器的CSS ID
            
        返回：
            str: CSS样式字符串
        """
        css_pos = self.calculate_chart_css_position(position)
        
        return f"""
        #{container_id} {{
            position: relative;
        }}
        .chart-overlay {{
            position: absolute;
            left: {css_pos.left:.1f}px;
            top: {css_pos.top:.1f}pt;
            width: {css_pos.width:.1f}px;
            height: {css_pos.height:.1f}pt;
            z-index: 10;
            pointer-events: auto;
        }}
        """
    
    def generate_chart_html_with_positioning(self, chart, chart_html: str) -> str:
        """
        为图表HTML添加定位信息。
        
        参数：
            chart: Chart对象
            chart_html: 原始图表HTML
            
        返回：
            str: 带定位的图表HTML
        """
        if not chart.position:
            # 没有定位信息，使用原始HTML
            return chart_html
            
        css_pos = self.calculate_chart_css_position(chart.position)
        
        # 添加内联样式进行精确定位
        positioned_html = f"""
        <div class="chart-overlay" style="
            position: absolute;
            left: {css_pos.left:.1f}px;
            top: {css_pos.top:.1f}pt;
            width: {css_pos.width:.1f}px;
            height: {css_pos.height:.1f}pt;
            z-index: 10;
        ">
            {chart_html}
        </div>
        """
        
        return positioned_html


def create_position_calculator(sheet: Sheet) -> ChartPositionCalculator:
    """
    工厂函数：创建图表定位计算器。
    
    参数：
        sheet: 工作表对象
        
    返回：
        ChartPositionCalculator: 计算器实例
    """
    return ChartPositionCalculator(sheet)