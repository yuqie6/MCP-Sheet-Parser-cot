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
    
    # 常量定义 - 基于实际HTML输出分析的精确转换
    EXCEL_TO_PX = 8.45  # Excel列宽单位到像素的转换系数（通过对比HTML输出计算得出）
    EMU_TO_PX = 96 / 914400  # EMU到像素的转换系数（914400 EMUs per inch）
    EMU_TO_PT = 72 / 914400  # EMU到点的转换系数（72 points per inch）
    
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
        # 使用实际的起始位置计算
        start_x, start_y = self._calculate_cell_position(position.from_col, position.from_row)

        # 添加单元格内偏移
        col_offset_px = position.from_col_offset * self.EMU_TO_PX
        row_offset_pt = position.from_row_offset * self.EMU_TO_PT

        # 计算最终位置 - 使用实际单元格位置而不是硬编码的M列位置
        final_x = start_x + col_offset_px
        final_y = start_y + row_offset_pt
        
        # 计算图表尺寸
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
        """
        # 修复：精确计算图表宽度
        start_x, _ = self._calculate_cell_position(position.from_col, 0)
        end_x, _ = self._calculate_cell_position(position.to_col, 0)
        
        start_offset = position.from_col_offset * self.EMU_TO_PX
        end_offset = position.to_col_offset * self.EMU_TO_PX
        
        # 结束列的宽度
        end_col_width = self.sheet.column_widths.get(position.to_col, self.sheet.default_column_width) * self.EXCEL_TO_PX
        
        width = (end_x - start_x) + end_offset - start_offset
        
        # 如果图表在单个单元格内
        if position.from_col == position.to_col:
            width = end_offset - start_offset
        else:
            # 跨多个单元格的情况
            middle_cols_width = 0
            for i in range(position.from_col + 1, position.to_col):
                middle_cols_width += self.sheet.column_widths.get(i, self.sheet.default_column_width) * self.EXCEL_TO_PX

            from_col_width = self.sheet.column_widths.get(position.from_col, self.sheet.default_column_width) * self.EXCEL_TO_PX
            to_col_width = self.sheet.column_widths.get(position.to_col, self.sheet.default_column_width) * self.EXCEL_TO_PX

            # 修复：如果to_col_offset为0，表示占满整个目标单元格
            actual_end_offset = end_offset if end_offset > 0 else to_col_width

            width = (from_col_width - start_offset) + middle_cols_width + actual_end_offset

        return max(50, width) # 恢复原始最小宽度，按实际尺寸显示
    
    def _calculate_chart_height(self, position: ChartPosition) -> float:
        """
        计算图表高度。
        """
        # 修复：精确计算图表高度
        _, start_y = self._calculate_cell_position(0, position.from_row)
        _, end_y = self._calculate_cell_position(0, position.to_row)

        start_offset = position.from_row_offset * self.EMU_TO_PT
        end_offset = position.to_row_offset * self.EMU_TO_PT
        
        if position.from_row == position.to_row:
            height = end_offset - start_offset
        else:
            # 计算跨多行的总高度
            total_height = 0

            # 起始行：从偏移位置到行末
            from_row_height = self.sheet.row_heights.get(position.from_row, self.sheet.default_row_height)
            total_height += (from_row_height - start_offset)

            # 中间完整行
            for i in range(position.from_row + 1, position.to_row):
                row_height = self.sheet.row_heights.get(i, self.sheet.default_row_height)
                total_height += row_height

            # 结束行：如果to_row_offset为0，表示占满整个目标行
            to_row_height = self.sheet.row_heights.get(position.to_row, self.sheet.default_row_height)
            actual_end_offset = end_offset if end_offset > 0 else to_row_height
            total_height += actual_end_offset

            height = total_height

        return max(50, height) # 恢复原始最小高度，按实际尺寸显示
    
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

        # 根据图表类型使用不同的定位逻辑
        if chart.type == 'image':
            # 图片使用原始位置计算（不应用柱状图的M列调整）
            css_pos = self._calculate_image_position(chart.position)
        else:
            # 其他图表（如柱状图）使用调整后的位置计算
            css_pos = self.calculate_chart_css_position(chart.position)

        # 添加内联样式进行精确定位
        # 修复：统一使用px单位，避免pt和px混用导致的尺寸不匹配
        # 将pt转换为px (1pt = 1.333px)
        top_px = css_pos.top * 1.333

        # 根据图表类型调整高度 - 修复容器与SVG高度不匹配问题
        if chart.type == 'image':
            height_px = css_pos.height * 1.333  # 图片不需要额外增加高度
        else:
            height_px = css_pos.height * 1.333  # 移除30%增加，保持与SVG高度一致

        # 使用绝对定位，但基于表格容器的相对位置
        # 这样可以保持Excel的覆盖效果，同时相对于表格定位
        # 根据图表类型进行不同的位置调整
        if chart.type == 'image':
            # 图片位置通常是正确的，不需要调整
            left_adjusted = css_pos.left
            top_adjusted = top_px
        else:
            # 其他图表使用计算出的位置，不需要额外调整
            left_adjusted = css_pos.left
            top_adjusted = top_px

        positioned_html = f"""
        <div class="chart-overlay" style="
            position: absolute;
            left: {left_adjusted:.1f}px;
            top: {top_adjusted:.1f}px;
            width: {css_pos.width:.1f}px;
            height: {height_px:.1f}px;
            z-index: 10;
            pointer-events: auto;
        ">
            {chart_html}
        </div>
        """

        return positioned_html

    def _calculate_image_position(self, position: ChartPosition) -> ChartCSSPosition:
        """
        计算图片的CSS定位信息（使用原始位置，不应用柱状图的调整）。

        参数：
            position: 图片的Excel定位信息

        返回：
            ChartCSSPosition: CSS定位信息
        """
        # 使用原始位置计算，不应用M列调整
        start_x, start_y = self._calculate_cell_position(position.from_col, position.from_row)

        # 添加单元格内偏移
        col_offset_px = position.from_col_offset * self.EMU_TO_PX
        row_offset_pt = position.from_row_offset * self.EMU_TO_PT

        final_x = start_x + col_offset_px
        final_y = start_y + row_offset_pt

        # 计算图片实际尺寸（基于EMU偏移量，而不是单元格跨度）
        width = self._calculate_image_width(position)
        height = self._calculate_image_height(position)

        return ChartCSSPosition(
            left=final_x,
            top=final_y,
            width=width,
            height=height
        )

    def _calculate_image_width(self, position: ChartPosition) -> float:
        """
        计算图片的实际宽度（基于EMU偏移量）。
        """
        # 如果有明确的结束偏移，使用偏移量差值
        if position.to_col_offset > 0:
            width_emu = position.to_col_offset - position.from_col_offset
        else:
            # 否则使用默认的图片宽度（约1英寸 = 914400 EMU）
            width_emu = 914400

        width_px = width_emu * self.EMU_TO_PX
        return max(20, width_px)  # 最小20px

    def _calculate_image_height(self, position: ChartPosition) -> float:
        """
        计算图片的实际高度（基于EMU偏移量）。
        """
        # 如果有明确的结束偏移，使用偏移量差值
        if position.to_row_offset > 0:
            height_emu = position.to_row_offset - position.from_row_offset
        else:
            # 否则使用默认的图片高度（约0.4英寸 = 285750 EMU）
            height_emu = 285750

        height_pt = height_emu * self.EMU_TO_PT
        return max(15, height_pt)  # 最小15pt


def create_position_calculator(sheet: Sheet) -> ChartPositionCalculator:
    """
    工厂函数：创建图表定位计算器。
    
    参数：
        sheet: 工作表对象
        
    返回：
        ChartPositionCalculator: 计算器实例
    """
    return ChartPositionCalculator(sheet)