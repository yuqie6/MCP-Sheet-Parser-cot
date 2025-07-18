"""
SVG图表渲染器：将Excel图表转换为高质量SVG格式。
"""

import math
from typing import Any
import xml.etree.ElementTree as ET
from xml.dom import minidom
from src.utils.color_utils import DEFAULT_CHART_COLORS, normalize_color, ensure_distinct_colors


class SVGChartRenderer:
    """Excel图表数据到SVG的渲染器。"""
    
    # 默认图表尺寸常量
    DEFAULT_WIDTH = 800
    DEFAULT_HEIGHT = 500

    def __init__(self, width: int = DEFAULT_WIDTH, height: int = DEFAULT_HEIGHT, show_axes: bool = False):
        """
        初始化SVG渲染器。
        参数：
            width: SVG宽度（像素）
            height: SVG高度（像素）
            show_axes: 是否显示坐标轴和网格线
        """
        # Excel默认图表大小15x7.5cm
        self.width = width
        self.height = height
        self.show_axes = show_axes
        # 如果不显示坐标轴，减少边距，整体往下移动
        if show_axes:
            self.margin = {'top': 70, 'right': 80, 'bottom': 60, 'left': 80}  # 增加top边距
        else:
            self.margin = {'top': 40, 'right': 20, 'bottom': 20, 'left': 20}  # 简洁模式，增加top边距
        self.plot_width = width - self.margin['left'] - self.margin['right']
        self.plot_height = height - self.margin['top'] - self.margin['bottom']
        
    def render_chart_to_svg(self, chart_data: dict[str, Any]) -> str:
        """
        渲染图表数据为SVG字符串。
        参数：
            chart_data 图表数据字典
        返回：
            SVG字符串
        """
        # 使用原始Excel尺寸如果可用
        if 'size' in chart_data:
            size_info = chart_data['size']
            if 'width_px' in size_info and 'height_px' in size_info:
                self.width = max(200, size_info['width_px'])  # 最小宽度200px
                self.height = max(150, size_info['height_px'])  # 最小高度150px
                # 重新计算内边距和绘图区域
                self.plot_width = self.width - self.margin['left'] - self.margin['right']
                self.plot_height = self.height - self.margin['top'] - self.margin['bottom']
        
        chart_type = chart_data.get('type', 'bar')
        
        if chart_type == 'bar':
            return self._render_bar_chart(chart_data)
        elif chart_type == 'line':
            return self._render_line_chart(chart_data)
        elif chart_type == 'pie':
            return self._render_pie_chart(chart_data)
        elif chart_type == 'area':
            return self._render_area_chart(chart_data)
        elif chart_type == 'image':
            return self._render_image_chart(chart_data)
        else:
            return self._render_fallback_chart(chart_data)

    def _render_legend_if_needed(self, svg: ET.Element, chart_data: dict[str, Any], series_list: list[dict], colors: list[str]) -> None:
        """渲染图例。"""
        show_legend = (
            chart_data.get('show_legend', False) or
            chart_data.get('legend', {}).get('show', False) or
            chart_data.get('legend', {}).get('enabled', False)
        )
        if show_legend:
            legend_style = chart_data.get('legend_style')
            # 如果有图例条目信息，使用条目信息；否则使用系列信息
            legend_entries = chart_data.get('legend', {}).get('entries', [])
            if legend_entries:
                # 使用图例条目信息创建图例
                legend_series = []
                for entry in legend_entries:
                    if not entry.get('delete', False):
                        legend_series.append({
                            'name': entry.get('text', f"Series {entry.get('index', 0) + 1}")
                        })
                self._draw_legend(svg, legend_series, colors, legend_style=legend_style)
            else:
                # 使用系列信息创建图例
                self._draw_legend(svg, series_list, colors, legend_style=legend_style)
    
    def _render_common_chart_elements(self, svg: ET.Element, chart_data: dict[str, Any], series_list: list[dict], colors: list[str]) -> None:
        """渲染通用图表元素（图例和注释）。"""
        self._render_legend_if_needed(svg, chart_data, series_list, colors)
        self._render_annotations(svg, chart_data)

    def _should_show_data_labels(self, chart_data: dict[str, Any], series_data: dict[str, Any] | None = None) -> bool:
        """判断是否显示数据标签。"""
        # 检查图表级别的数据标签设置
        if chart_data.get('show_data_labels', False):
            return True
            
        # 检查系列级别的数据标签设置
        if series_data and series_data.get('show_data_labels', False):
            return True
            
        # 检查嵌套的数据标签配置
        data_labels_config = chart_data.get('data_labels', {})
        if isinstance(data_labels_config, dict) and data_labels_config.get('show', False):
            return True
            
        return False
    
    def _create_data_label_element(self, svg: ET.Element, x: float, y: float, text: str, 
                                   font_size: str = '10px', fill: str = 'white', 
                                   text_anchor: str = 'middle') -> None:
        """创建数据标签元素。"""
        text_elem = ET.SubElement(svg, 'text', {
            'x': str(x),
            'y': str(y),
            'text-anchor': text_anchor,
            'alignment-baseline': 'middle',
            'class': 'axis-label',
            'fill': fill,
            'font-size': font_size
        })
        text_elem.text = text
    
    def _handle_chart_error(self, chart_type: str, error: Exception) -> str:
        """处理渲染错误，返回错误占位符。"""
        svg = self._create_svg_root(f"Chart Error: {chart_type}")
        
        # 绘制错误占位符
        ET.SubElement(svg, 'rect', {
            'x': str(self.margin['left']),
            'y': str(self.margin['top']),
            'width': str(self.plot_width),
            'height': str(self.plot_height),
            'fill': '#ffebee',
            'stroke': '#f44336',
            'stroke-dasharray': '5,5'
        })
        
        text = ET.SubElement(svg, 'text', {
            'x': str(self.width // 2),
            'y': str(self.height // 2),
            'text-anchor': 'middle',
            'class': 'axis-label',
            'fill': '#d32f2f'
        })
        text.text = f"Error rendering {chart_type}: {str(error)}"
        
        return self._format_svg(svg)

    def _create_svg_root(self, title: str = "", title_style: dict[str, Any] | None = None) -> ET.Element:
        """
        创建SVG根元素。
        参数：title 图表标题，title_style 标题样式
        """
        svg = ET.Element('svg', {
            'width': f'{self.width}px',
            'height': f'{self.height}px',
            'viewBox': f'0 0 {self.width} {self.height}',
            'xmlns': 'http://www.w3.org/2000/svg',
            'class': 'excel-chart-svg'
        })
        
        # 添加样式定义
        style = ET.SubElement(svg, 'style')
        style.text = self._get_chart_css()
        
        # 添加标题
        if title:
            title_attrs = {
                'x': str(self.width // 2),
                'y': '35',  # 往下移动10px
                'text-anchor': 'middle',
                'class': 'chart-title'
            }
            
            # 合并外部传入的样式
            if title_style:
                style_str = ""
                if 'font_family' in title_style and title_style['font_family']:
                    style_str += f"font-family: '{title_style['font_family']}';"
                if 'font_size' in title_style and title_style['font_size']:
                    style_str += f"font-size: {title_style['font_size']}px;"
                if 'color' in title_style and title_style['color']:
                    style_str += f"fill: {title_style['color']};"
                if 'bold' in title_style and title_style['bold']:
                    style_str += "font-weight: bold;"
                if style_str:
                    title_attrs['style'] = style_str
            
            title_elem = ET.SubElement(svg, 'text', title_attrs)
            title_elem.text = title
        
        return svg

    def _get_series_colors(self, chart_data: dict[str, Any]) -> list[str]:
        """
        获取系列颜色，优先使用Excel原色。
        """
        if 'colors' in chart_data and chart_data['colors']:
            return chart_data['colors']
        
        colors = []
        for series in chart_data.get('series', []):
            if 'color' in series and series['color']:
                colors.append(series['color'])
        
        if colors:
            return colors
        
        return DEFAULT_CHART_COLORS
    
    def _get_chart_css(self) -> str:
        """获取图表CSS样式。"""
        return """
        .excel-chart-svg {
            font-family: 'Microsoft YaHei', 'SimHei', 'PingFang SC', 'Hiragino Sans GB', 'Source Han Sans SC', 'Noto Sans CJK SC', 'Segoe UI', Arial, sans-serif;
        }
        .chart-title {
            font-size: 18px;  /* 基于搜索结果：Excel默认18pt */
            font-weight: bold;
            fill: #333;
        }
        .axis-title {
            font-size: 12px;
            fill: #666;
        }
        .axis-label {
            font-size: 11px;
            fill: #666;
        }
        .grid-line {
            stroke: #e0e0e0;
            stroke-width: 0.5;
        }
        .axis-line {
            stroke: #666;
            stroke-width: 1;
        }
        .bar-rect {
            fill-opacity: 0.8;
            stroke-width: 1;
            stroke: #fff;
        }
        .bar-rect:hover {
            fill-opacity: 1;
        }
        .line-path {
            fill: none;
            stroke-width: 2;
        }
        .line-point {
            r: 3;
            fill-opacity: 0.8;
        }
        .line-point:hover {
            r: 4;
        }
        .area-path {
            fill-opacity: 0.3;
            stroke-width: 2;
        }
        .area-point {
            r: 3;
            fill-opacity: 0.8;
        }
        .area-point:hover {
            r: 4;
        }
        .pie-slice {
            stroke: #fff;
            stroke-width: 1;
        }
        .pie-slice:hover {
            stroke-width: 2;
        }
        .legend-item {
            font-size: 11px;
            fill: #666;
        }
        """
    
    def _render_bar_chart(self, chart_data: dict[str, Any]) -> str:
        """渲染柱状图。"""
        title_style = chart_data.get('title_style')
        svg = self._create_svg_root(chart_data.get('title', ''), title_style=title_style)
        series_list = chart_data.get('series', [])

        # 智能判断是否应该去重X轴标签
        should_deduplicate = self._should_deduplicate_x_labels(series_list)

        # 保存当前系列，供其他方法使用
        self.current_series = series_list

        # 确保颜色列表与系列列表匹配
        colors = self._get_series_colors(chart_data)
        if len(colors) < len(series_list):
            colors.extend(DEFAULT_CHART_COLORS * (len(series_list) - len(colors)))

        if not series_list:
            return self._format_svg(svg)

        all_x_labels = []
        all_y_values = []
        for series in series_list:
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])
            all_x_labels.extend(x_data)
            all_y_values.extend(y_data)

        if not all_x_labels or not all_y_values:
            return self._format_svg(svg)

        # 根据判断结果决定X轴标签策略
        if should_deduplicate:
            # 分组柱状图：去重X轴标签
            unique_x_labels = list(dict.fromkeys(all_x_labels))
            display_x_labels = unique_x_labels
        else:
            # 连续柱状图：保留所有X轴标签
            unique_x_labels = list(dict.fromkeys(all_x_labels))  # 仍需要用于计算范围
            display_x_labels = all_x_labels
        # 修复：支持负数值，Y轴范围应该包含所有数据
        y_min = min(all_y_values) if all_y_values else 0

        # 确保y_max在传递给float()之前绝不为None
        y_val = chart_data.get('y_axis_max')
        if y_val is None or not isinstance(y_val, (int, float)):
            y_val = max(all_y_values) if all_y_values else 1

        # 确保y_min和y_max是数值类型，并处理可能的转换错误
        try:
            y_min = float(y_min)
            y_max = float(y_val)
        except (ValueError, TypeError):
            y_min = 0.0
            y_max = 1.0

        # 如果所有值都是正数，从0开始；如果有负数，包含负数范围
        if y_min > 0:
            y_min = 0  # 正数图表从0开始更直观

        y_range = y_max - y_min if y_max != y_min else 1

        # 绘制X轴标签（保持Excel样式）
        self._draw_x_axis_labels(svg, display_x_labels)

        # 如果有负数，绘制零线
        if y_min < 0 and y_range > 0:
            zero_line_y = self.margin['top'] + self.plot_height - ((-y_min) / y_range) * self.plot_height
            zero_line = ET.SubElement(svg, 'line', {
                'x1': str(self.margin['left']),
                'y1': str(zero_line_y),
                'x2': str(self.margin['left'] + self.plot_width),
                'y2': str(zero_line_y),
                'stroke': '#666',
                'stroke-width': '1',
                'class': 'zero-line'
            })

        colors = self._get_series_colors(chart_data)

        # 根据标签策略决定柱子渲染方式
        if should_deduplicate:
            # 分组柱状图：每个标签下有多个系列的柱子
            num_groups = len(unique_x_labels)
            num_series = len(series_list)
            group_width = self.plot_width / num_groups
            bar_width = group_width / num_series * 0.8  # 每个系列的柱子宽度
        else:
            # 连续柱状图：每个数据点都是独立的柱子
            total_bars = len(all_x_labels)
            bar_group_width = self.plot_width / total_bars
            bar_width = bar_group_width * 0.5

        bar_index = 0  # 全局柱子索引（用于连续模式）

        for series_idx, series in enumerate(series_list):
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])

            # 收集数据标签位置信息
            x_positions = []
            y_positions = []

            for i, (x_label, y_value) in enumerate(zip(x_data, y_data)):
                color = normalize_color(series.get('color') or colors[series_idx % len(colors)])

                # 根据标签策略计算柱子位置
                if should_deduplicate:
                    # 分组柱状图：每个标签下有多个系列的柱子
                    group_index = unique_x_labels.index(x_label) if x_label in unique_x_labels else i
                    group_center = self.margin['left'] + (group_index + 0.5) * group_width
                    bar_offset = (series_idx - (num_series - 1) / 2) * bar_width
                    bar_center_x = group_center + bar_offset
                    bar_x = bar_center_x - bar_width / 2
                else:
                    # 连续柱状图：每个数据点都是独立的柱子
                    bar_center_x = self.margin['left'] + (bar_index + 0.5) * bar_group_width
                    bar_x = bar_center_x - bar_width / 2

                # 计算柱子高度和位置，支持负数
                if y_range > 0:
                    # 计算零线位置（基准线）
                    zero_line_y = self.margin['top'] + self.plot_height - ((-y_min) / y_range) * self.plot_height

                    # 计算柱子的顶部和底部位置
                    value_y = self.margin['top'] + self.plot_height - ((y_value - y_min) / y_range) * self.plot_height

                    if y_value >= 0:
                        # 正数：从零线向上
                        bar_y = value_y
                        bar_height = zero_line_y - value_y
                    else:
                        # 负数：从零线向下
                        bar_y = zero_line_y
                        bar_height = value_y - zero_line_y

                    # 确保柱子至少有最小高度可见
                    bar_height = max(2, abs(bar_height))
                else:
                    bar_height = 5
                    bar_y = self.margin['top'] + self.plot_height - bar_height

                # 绘制柱子
                rect = ET.SubElement(svg, 'rect', {
                    'x': str(bar_x),
                    'y': str(bar_y),
                    'width': str(bar_width),
                    'height': str(bar_height),
                    'fill': color,
                    'class': 'bar-rect'
                })

                # 收集位置信息用于数据标签
                x_positions.append(bar_center_x)
                y_positions.append(bar_y)

                # 使用新的通用方法显示数据标签，支持正负数
                if self._should_show_data_labels(chart_data, series) and bar_height > 8:
                    # 根据正负值调整标签位置
                    if y_value >= 0:
                        # 正数：标签在柱子顶部上方
                        label_y = bar_y - 5
                    else:
                        # 负数：标签在柱子底部下方
                        label_y = bar_y + bar_height + 15

                    self._create_data_label_element(
                        svg,
                        bar_x + bar_width / 2,
                        label_y,
                        str(int(y_value))
                    )

                # 只在连续模式下更新bar_index
                if not should_deduplicate:
                    bar_index += 1

            # 渲染该系列的数据标签
            if x_positions and y_positions:
                self._render_data_labels(svg, series, x_positions, y_positions, colors)

        # 渲染图例和注释
        self._render_common_chart_elements(svg, chart_data, series_list, colors)
        
        return self._format_svg(svg)

    def _render_line_chart(self, chart_data: dict[str, Any]) -> str:
        """渲染折线图。"""
        title_style = chart_data.get('title_style')
        svg = self._create_svg_root(chart_data.get('title', ''), title_style=title_style)
        series_list = chart_data.get('series', [])
        
        if not series_list:
            return self._format_svg(svg)
        
        # 获取所有数据点
        all_x_labels = []
        all_y_values = []
        for series in series_list:
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])
            all_x_labels.extend(x_data)
            all_y_values.extend(y_data)
        
        if not all_x_labels or not all_y_values:
            return self._format_svg(svg)
        
        unique_x_labels = list(dict.fromkeys(all_x_labels))
        y_min = min(all_y_values) if all_y_values else 0
        y_max = max(all_y_values) if all_y_values else 1
        y_range = y_max - y_min if y_max != y_min else 1
        
        # 绘制坐标轴
        self._draw_axes(svg, unique_x_labels, y_min, y_max)
        
        # 绘制数据系列
        colors = self._get_series_colors(chart_data)
        
        for series_idx, series in enumerate(series_list):
            color = normalize_color(colors[series_idx % len(colors)])
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])
            
            if len(x_data) < 2 or len(y_data) < 2:
                continue
            
            # 创建路径点
            points = []
            for x_label, y_value in zip(x_data, y_data):
                if x_label in unique_x_labels:
                    x_pos = unique_x_labels.index(x_label)                    
                    if len(unique_x_labels) == 1:
                        x_coord = self.margin['left'] + self.plot_width / 2
                    else:
                        x_coord = self.margin['left'] + x_pos * (self.plot_width / (len(unique_x_labels) - 1))
                    y_coord = self.margin['top'] + self.plot_height - (y_value - y_min) / y_range * self.plot_height
                    points.append((x_coord, y_coord))
            
            if len(points) >= 2:
                # 创建路径
                path_data = f"M {points[0][0]} {points[0][1]}"
                for x, y in points[1:]:
                    path_data += f" L {x} {y}"
                
                path = ET.SubElement(svg, 'path', {
                    'd': path_data,
                    'stroke': color,
                    'class': 'line-path'
                })
                
                # 添加数据点
                for x, y in points:
                    circle = ET.SubElement(svg, 'circle', {
                        'cx': str(x),
                        'cy': str(y),
                        'r': '3',  # 添加半径属性
                        'fill': color,
                        'class': 'line-point'
                    })
        
        # 渲染图例和注释
        self._render_common_chart_elements(svg, chart_data, series_list, colors)
        
        return self._format_svg(svg)
    
    def _render_pie_chart(self, chart_data: dict[str, Any]) -> str:
        """渲染饼图。"""
        # 设置饼图标志，供图例位置计算使用
        self._is_pie_chart = True

        # 饼图标题：优先使用图表标题，如果没有则使用系列名称
        title = chart_data.get('title', '')
        if not title or title.startswith('Chart '):
            series_list = chart_data.get('series', [])
            if series_list and series_list[0].get('name'):
                title = series_list[0]['name']

        # 饼图需要特殊的标题位置处理，先创建不带标题的SVG
        svg = self._create_svg_root("")
        series_list = chart_data.get('series', [])
        
        if not series_list or not series_list[0].get('y_data'):
            return self._format_svg(svg)
        
        series = series_list[0]  # 饼图通常只有一个系列
        labels = series.get('x_data', [])
        values = series.get('y_data', [])
        
        if not values:
            return self._format_svg(svg)
        
        # 计算饼图参数
        total = sum(values)
        if total == 0:
            return self._format_svg(svg)
        
        # 为图例预留空间，饼图偏左放置
        legend_width = 100  # 图例区域宽度
        pie_area_width = self.width - legend_width
        center_x = pie_area_width // 2
        # 调整垂直位置，考虑标题空间，避免过于偏上
        title_height = 40 if chart_data.get('title') else 20
        center_y = title_height + (self.height - title_height) // 2
        radius = min(pie_area_width, self.height - title_height) // 2 - 30
        
        # 优先使用从Excel提取的颜色
        colors = []
        if series.get('colors'):  # 饼图的每个片段可能有不同颜色
            colors = series['colors']
        elif 'colors' in chart_data and chart_data['colors']:
            colors = chart_data['colors']
        else:
            colors = DEFAULT_CHART_COLORS + ['#A5A5A5', '#70E000']  # 扩展颜色用于多系列饼图

        # 确保有足够的不重复颜色用于所有片段
        colors = ensure_distinct_colors(colors, len(values))
        
        current_angle = 0
        for i, (label, value) in enumerate(zip(labels, values)):
            angle = (value / total) * 360
            color = normalize_color(colors[i % len(colors)])
            
            # 计算扇形路径
            start_angle_rad = current_angle * 3.14159 / 180
            end_angle_rad = (current_angle + angle) * 3.14159 / 180
            
            start_x = center_x + radius * math.cos(start_angle_rad)
            start_y = center_y + radius * math.sin(start_angle_rad)
            end_x = center_x + radius * math.cos(end_angle_rad)
            end_y = center_y + radius * math.sin(end_angle_rad)
            
            large_arc = 1 if angle > 180 else 0
            
            path_data = f"M {center_x} {center_y} L {start_x} {start_y} A {radius} {radius} 0 {large_arc} 1 {end_x} {end_y} Z"
            
            path = ET.SubElement(svg, 'path', {
                'd': path_data,
                'fill': color,
                'class': 'pie-slice'
            })
            
            # 添加数据标签（如果启用）
            series_data = series_list[0] if series_list else {}
            data_labels = series_data.get('data_labels', {})

            if data_labels.get('enabled', True):  # 饼图默认显示标签
                mid_angle = (current_angle + angle / 2) * 3.14159 / 180
                label_x = center_x + (radius + 20) * math.cos(mid_angle)
                label_y = center_y + (radius + 20) * math.sin(mid_angle)

                # 构建标签文本
                label_text = ""

                if data_labels.get('show_category_name', True):  # 饼图默认显示分类名
                    label_text = label

                if data_labels.get('show_value', False):
                    if label_text:
                        label_text += f": {value}"
                    else:
                        label_text = str(value)

                if data_labels.get('show_percent', True):  # 饼图默认显示百分比
                    percentage = (value / total) * 100
                    if label_text:
                        label_text += f" ({percentage:.1f}%)"
                    else:
                        label_text = f"{percentage:.1f}%"

                # 如果没有设置任何显示选项，默认显示标签和百分比
                if not label_text:
                    percentage = (value / total) * 100
                    label_text = f"{label} ({percentage:.1f}%)"

                text = ET.SubElement(svg, 'text', {
                    'x': str(label_x),
                    'y': str(label_y),
                    'text-anchor': 'middle',
                    'class': 'axis-label'
                })
                text.text = label_text
            
            current_angle += angle

        # 渲染图例和注释 - 饼图使用分类标签作为图例
        pie_legend_series = []
        for i, label in enumerate(labels):
            pie_legend_series.append({'name': label})

        # 使用分类标签创建图例，而不是系列名称
        show_legend = chart_data.get('legend', {}).get('enabled', True)
        if show_legend:
            legend_style = chart_data.get('legend_style')
            self._draw_legend(svg, pie_legend_series, colors, legend_style=legend_style)

        # 添加饼图标题（在饼图上方）
        if title:
            title_y = 20  # 标题位置在饼图上方
            title_attrs = {
                'x': str(self.width // 2),
                'y': str(title_y),
                'text-anchor': 'middle',
                'class': 'chart-title'
            }
            title_elem = ET.SubElement(svg, 'text', title_attrs)
            title_elem.text = title

        # 渲染注释
        self._render_annotations(svg, chart_data)

        # 清除饼图标志
        self._is_pie_chart = False

        return self._format_svg(svg)

    def _render_area_chart(self, chart_data: dict[str, Any]) -> str:
        """渲染面积图。"""
        svg = self._create_svg_root(chart_data.get('title', ''))
        series_list = chart_data.get('series', [])

        if not series_list:
            return self._format_svg(svg)
        
        # 获取所有数据点
        all_x_labels = []
        all_y_values = []
        for series in series_list:
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])
            all_x_labels.extend(x_data)
            all_y_values.extend(y_data)
        
        if not all_x_labels or not all_y_values:
            return self._format_svg(svg)
        
        unique_x_labels = list(dict.fromkeys(all_x_labels))
        y_min = min(all_y_values) if all_y_values else 0
        y_max = max(all_y_values) if all_y_values else 1
        y_range = y_max - y_min if y_max != y_min else 1
        
        # 绘制坐标轴
        self._draw_axes(svg, unique_x_labels, y_min, y_max)
        
        # 绘制数据系列
        colors = self._get_series_colors(chart_data)
        baseline_y = self.margin['top'] + self.plot_height  # 基线位置
        
        for series_idx, series in enumerate(series_list):
            color = normalize_color(colors[series_idx % len(colors)])
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])
            
            if len(x_data) < 2 or len(y_data) < 2:
                continue
            
            # 创建面积路径点
            points = []
            for x_label, y_value in zip(x_data, y_data):
                if x_label in unique_x_labels:
                    x_pos = unique_x_labels.index(x_label)
                    if len(unique_x_labels) == 1:
                        x_coord = self.margin['left'] + self.plot_width / 2
                    else:
                        x_coord = self.margin['left'] + x_pos * (self.plot_width / (len(unique_x_labels) - 1))
                    y_coord = self.margin['top'] + self.plot_height - (y_value - y_min) / y_range * self.plot_height
                    points.append((x_coord, y_coord))
            
            if len(points) >= 2:
                # 创建面积路径（从基线开始，到数据点，再回到基线）
                path_data = f"M {points[0][0]} {baseline_y}"  # 从基线开始
                path_data += f" L {points[0][0]} {points[0][1]}"  # 到第一个数据点
                
                # 连接所有数据点
                for x, y in points[1:]:
                    path_data += f" L {x} {y}"
                
                # 回到基线闭合路径
                path_data += f" L {points[-1][0]} {baseline_y} Z"
                
                # 绘制填充区域
                area_path = ET.SubElement(svg, 'path', {
                    'd': path_data,
                    'fill': color,
                    'fill-opacity': '0.3',
                    'stroke': color,
                    'stroke-width': '2',
                    'class': 'area-path'
                })
                
                # 添加数据点
                for x, y in points:
                    circle = ET.SubElement(svg, 'circle', {
                        'cx': str(x),
                        'cy': str(y),
                        'r': '3',
                        'fill': color,
                        'class': 'area-point'
                    })
        
        # 渲染图例和注释
        self._render_common_chart_elements(svg, chart_data, series_list, colors)
        
        return self._format_svg(svg)
    
    def _render_image_chart(self, chart_data: dict[str, Any]) -> str:
        """渲染图像（支持图片和占位符）。"""
        # 检查是否有有效的图片数据 - 支持多种数据结构
        img_data = None
        
        # 尝试从不同的位置获取图片数据
        if isinstance(chart_data, dict):
            # 方式1：直接从chart_data中获取
            img_data = chart_data.get('image_data')
            
            # 方式2：从嵌套字典中获取（这是xlsx_parser存储的方式）
            if not img_data and 'image_data' in chart_data:
                nested_data = chart_data['image_data']
                if isinstance(nested_data, dict):
                    img_data = nested_data.get('image_data')
                else:
                    img_data = nested_data
        
        if img_data and isinstance(img_data, bytes) and len(img_data) > 10:
            # 有真实图片数据，渲染为HTML img标签而不是SVG
            import base64

            # 检测图片格式
            img_format = 'png'  # 默认
            if img_data.startswith(b'\x89PNG'):
                img_format = 'png'
            elif img_data.startswith(b'\xff\xd8\xff'):
                img_format = 'jpeg'
            elif img_data.startswith(b'GIF'):
                img_format = 'gif'

            # 转换为Base64
            img_base64 = base64.b64encode(img_data).decode('utf-8')
            data_url = f"data:image/{img_format};base64,{img_base64}"

            # 返回HTML img标签 - 正常显示，无背景
            return f'''
            <div class="chart-svg-wrapper" style="background: none; padding: 0; min-height: auto;">
                <img src="{data_url}" alt="Excel Image" style="max-width: 100%; height: auto; border-radius: 4px; background: transparent;" />
            </div>
            '''

        # 没有有效图片数据，渲染占位符SVG
        svg = self._create_svg_root(chart_data.get('title', 'Image'))

        # 绘制图像占位符（移除不必要的背景框）
        placeholder_rect = ET.SubElement(svg, 'rect', {
            'x': str(self.margin['left']),
            'y': str(self.margin['top']),
            'width': str(self.plot_width),
            'height': str(self.plot_height),
            'fill': 'none',
            'stroke': 'none'
        })
        
        # 添加图像图标（简化的SVG图标）
        icon_size = min(self.plot_width, self.plot_height) * 0.3
        icon_x = self.margin['left'] + (self.plot_width - icon_size) / 2
        icon_y = self.margin['top'] + (self.plot_height - icon_size) / 2
        
        # 图像图标框（简化样式）
        icon_rect = ET.SubElement(svg, 'rect', {
            'x': str(icon_x),
            'y': str(icon_y),
            'width': str(icon_size),
            'height': str(icon_size * 0.7),
            'fill': 'none',
            'stroke': '#9ca3af',
            'stroke-width': '1.5',
            'rx': '6'
        })
        
        # 图像图标中的"山"形状（优化颜色）
        mountain_points = f"{icon_x + icon_size * 0.1},{icon_y + icon_size * 0.5} {icon_x + icon_size * 0.3},{icon_y + icon_size * 0.3} {icon_x + icon_size * 0.5},{icon_y + icon_size * 0.45} {icon_x + icon_size * 0.7},{icon_y + icon_size * 0.25} {icon_x + icon_size * 0.9},{icon_y + icon_size * 0.5}"
        mountain = ET.SubElement(svg, 'polyline', {
            'points': mountain_points,
            'fill': 'none',
            'stroke': '#9ca3af',
            'stroke-width': '1.5',
            'stroke-linejoin': 'round'
        })
        
        # 图像图标中的"太阳"（优化颜色和大小）
        sun_cx = icon_x + icon_size * 0.75
        sun_cy = icon_y + icon_size * 0.2
        sun = ET.SubElement(svg, 'circle', {
            'cx': str(sun_cx),
            'cy': str(sun_cy),
            'r': str(icon_size * 0.06),
            'fill': '#f59e0b'
        })
        
        # 添加文本说明（优化颜色）
        text_y = icon_y + icon_size * 0.85
        text = ET.SubElement(svg, 'text', {
            'x': str(self.width // 2),
            'y': str(text_y),
            'text-anchor': 'middle',
            'class': 'axis-label',
            'fill': '#9ca3af'
        })
        text.text = "图像内容"
        
        # 如果有图像数据，显示额外信息（优化颜色）
        if img_data:
            info_text = ET.SubElement(svg, 'text', {
                'x': str(self.width // 2),
                'y': str(text_y + 20),
                'text-anchor': 'middle',
                'class': 'axis-label',
                'fill': '#9ca3af',
                'font-size': '10'
            })
            info_text.text = "包含嵌入图像数据"
        
        return self._format_svg(svg)

    def _render_fallback_chart(self, chart_data: dict[str, Any]) -> str:
        """渲染不支持类型的占位图表。"""
        svg = self._create_svg_root(chart_data.get('title', 'Unsupported Chart'))
        
        # 绘制占位符
        rect = ET.SubElement(svg, 'rect', {
            'x': str(self.margin['left']),
            'y': str(self.margin['top']),
            'width': str(self.plot_width),
            'height': str(self.plot_height),
            'fill': '#f0f0f0',
            'stroke': '#ccc',
            'stroke-dasharray': '5,5'
        })
        
        text = ET.SubElement(svg, 'text', {
            'x': str(self.width // 2),
            'y': str(self.height // 2),
            'text-anchor': 'middle',
            'class': 'axis-label'
        })
        text.text = f"Chart type '{chart_data.get('type', 'unknown')}' not supported"
        
        return self._format_svg(svg)

    def _draw_x_axis_labels(self, svg: ET.Element, x_labels: list[str]):
        """绘制X轴标签。"""
        # 使用传入的唯一标签，而不是重复的标签
        unique_labels = x_labels

        # 计算标签组的宽度（每个标签对应多个系列的柱子）
        if not unique_labels:
            return

        label_group_width = self.plot_width / len(unique_labels)
        series_count = len(self.current_series)

        # X轴标签 - 每个标签居中于对应的柱子组
        for i, label in enumerate(unique_labels):
            # 计算标签位置（居中于柱子组）
            group_center = self.margin['left'] + (i + 0.5) * label_group_width

            text = ET.SubElement(svg, 'text', {
                'x': str(group_center),
                'y': str(self.margin['top'] + self.plot_height + 15),
                'text-anchor': 'middle',
                'class': 'axis-label',
                'fill': '#666',  # 确保标签可见
                'font-size': '10px'  # 稍微减小字体，避免重叠
            })
            text.text = str(label)

    def _draw_axes(self, svg: ET.Element, x_labels: list[str], y_min: float, y_max: float):
        """绘制坐标轴、网格线和标签。"""
        # X轴
        x_axis = ET.SubElement(svg, 'line', {
            'x1': str(self.margin['left']),
            'y1': str(self.margin['top'] + self.plot_height),
            'x2': str(self.margin['left'] + self.plot_width),
            'y2': str(self.margin['top'] + self.plot_height),
            'class': 'axis-line'
        })

        # Y轴
        y_axis = ET.SubElement(svg, 'line', {
            'x1': str(self.margin['left']),
            'y1': str(self.margin['top']),
            'x2': str(self.margin['left']),
            'y2': str(self.margin['top'] + self.plot_height),
            'class': 'axis-line'
        })

        # 复用X轴标签绘制
        self._draw_x_axis_labels(svg, x_labels)
        
        # Y轴标签
        y_range = y_max - y_min if y_max != y_min else 1
        tick_count = 5
        for i in range(tick_count + 1):
            y_value = y_min + i * y_range / tick_count
            y_pos = self.margin['top'] + self.plot_height - i * (self.plot_height / tick_count)
            
            # 网格线
            grid_line = ET.SubElement(svg, 'line', {
                'x1': str(self.margin['left']),
                'y1': str(y_pos),
                'x2': str(self.margin['left'] + self.plot_width),
                'y2': str(y_pos),
                'class': 'grid-line'
            })
            
            # 标签
            text = ET.SubElement(svg, 'text', {
                'x': str(self.margin['left'] - 10),
                'y': str(y_pos),
                'text-anchor': 'end',
                'alignment-baseline': 'middle',
                'class': 'axis-label'
            })
            text.text = f"{y_value:.1f}"

    def _draw_legend(self, svg: ET.Element, series_list: list[dict], colors: list[str], legend_style: dict[str, Any] | None = None):
        """
        绘制图例。
        参数：svg根元素，series_list系列列表，colors颜色列表，legend_style字体样式
        """
        # 图例位置计算：根据图表类型调整
        if hasattr(self, '_is_pie_chart') and self._is_pie_chart:
            # 饼图：图例放在右侧
            legend_width = 100  # 与饼图布局中的legend_width保持一致
            pie_area_width = self.width - legend_width
            legend_x = pie_area_width + 10  # 饼图区域右侧，留10px间距
            legend_y = self.height // 2 - (len(series_list) * 20) // 2  # 垂直居中
        else:
            # 其他图表：图例放在右上角
            legend_text_width = 80
            legend_x = max(10, self.width - legend_text_width - 20)
            legend_y = self.margin['top']
        
        # 准备基础样式
        base_style = {'class': 'legend-item'}
        
        # 合并外部传入的样式
        if legend_style:
            style_str = ""
            if 'font_family' in legend_style and legend_style['font_family']:
                style_str += f"font-family: '{legend_style['font_family']}';"
            if 'font_size' in legend_style and legend_style['font_size']:
                style_str += f"font-size: {legend_style['font_size']}px;"
            if 'color' in legend_style and legend_style['color']:
                style_str += f"fill: {legend_style['color']};"
            if style_str:
                base_style['style'] = style_str

        for i, series in enumerate(series_list):
            color = normalize_color(colors[i % len(colors)])
            series_name = series.get('name', f'Series {i + 1}')

            # 图例色块
            ET.SubElement(svg, 'rect', {
                'x': str(legend_x),
                'y': str(legend_y + i * 20),
                'width': '12',
                'height': '12',
                'fill': color
            })

            # 图例文字的属性
            text_attrs = {
                'x': str(legend_x + 16),
                'y': str(legend_y + i * 20 + 9),
                **base_style
            }

            text = ET.SubElement(svg, 'text', text_attrs)
            text.text = series_name

    def _render_data_labels(self, svg: ET.Element, series_data: dict, x_positions: list, y_positions: list, colors: list | None = None) -> None:
        """
        渲染数据标签。
        """
        data_labels = series_data.get('data_labels', {})
        if not data_labels.get('enabled', False):
            return

        y_data = series_data.get('y_data', [])
        x_data = series_data.get('x_data', [])

        for i, (x_pos, y_pos) in enumerate(zip(x_positions, y_positions)):
            if i >= len(y_data):
                break

            label_text = ""

            # 根据设置决定显示什么内容
            if data_labels.get('show_value', False) and i < len(y_data):
                label_text = str(y_data[i])

            if data_labels.get('show_category_name', False) and i < len(x_data):
                if label_text:
                    label_text += f" ({x_data[i]})"
                else:
                    label_text = str(x_data[i])

            if data_labels.get('show_series_name', False):
                series_name = series_data.get('name', '')
                if series_name:
                    if label_text:
                        label_text = f"{series_name}: {label_text}"
                    else:
                        label_text = series_name

            if data_labels.get('show_percent', False):
                # 计算百分比
                if 'total' in series_data:
                    # 使用预计算的总数（主要用于饼图）
                    total = series_data['total']
                else:
                    # 计算当前系列的总数
                    total = sum(y_data) if y_data else 0

                if total > 0 and i < len(y_data):
                    percent = (y_data[i] / total) * 100
                    if label_text:
                        label_text += f" ({percent:.1f}%)"
                    else:
                        label_text = f"{percent:.1f}%"

            # 如果没有设置任何显示选项，但数据标签已启用，默认显示值
            if not label_text and data_labels.get('enabled', False) and i < len(y_data):
                # 检查是否有任何显示选项被设置
                has_any_show_option = any([
                    data_labels.get('show_value', False),
                    data_labels.get('show_category_name', False),
                    data_labels.get('show_series_name', False),
                    data_labels.get('show_percent', False)
                ])

                # 如果没有明确的显示选项，默认显示值
                if not has_any_show_option:
                    label_text = str(y_data[i])

            if label_text:
                # 调整标签位置，避免与数据点重叠
                label_y = y_pos - 5  # 标签显示在数据点上方

                text_elem = ET.SubElement(svg, 'text', {
                    'x': str(x_pos),
                    'y': str(label_y),
                    'text-anchor': 'middle',
                    'font-family': 'Arial, sans-serif',
                    'font-size': '10',
                    'fill': '#333'
                })
                text_elem.text = label_text

    def _render_annotations(self, svg: ET.Element, chart_data: dict) -> None:
        """
        渲染图表注释（不含标题）。
        """
        annotations = chart_data.get('annotations', [])
        if not annotations:
            return

        for annotation in annotations:
            annotation_type = annotation.get('type', '')
            text = annotation.get('text', '')
            position = annotation.get('position', '')

            if not text:
                continue

            # 跳过标题类型的注释，因为标题已经在_create_svg_root中处理了
            if annotation_type == 'title':
                continue

            # 根据注释类型和位置确定坐标
            x, y = self._get_annotation_position(annotation_type, position)

            text_elem = ET.SubElement(svg, 'text', {
                'x': str(x),
                'y': str(y),
                'text-anchor': 'middle',
                'font-family': 'Arial, sans-serif',
                'font-size': '11',
                'font-weight': 'normal',
                'fill': '#666'
            })
            text_elem.text = text

    def _get_annotation_position(self, annotation_type: str, position: str) -> tuple[int, int]:
        """
        获取注释坐标。
        """
        if annotation_type == 'title' and position == 'top':
            return (self.width // 2, 15)
        elif annotation_type == 'axis_title':
            if position == 'bottom':
                return (self.width // 2, self.height - 5)
            elif position == 'left':
                return (15, self.height // 2)
            elif position == 'right':
                return (self.width - 15, self.height // 2)

        # 默认位置
        return (self.width // 2, self.height // 2)

    def _format_svg(self, svg: ET.Element) -> str:
        """格式化SVG为字符串。"""
        rough_string = ET.tostring(svg, encoding='unicode')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")[22:]  # 去掉XML声明

    def _should_deduplicate_x_labels(self, series_list: list) -> bool:
        """
        判断是否去重X轴标签。
        """
        if not series_list or len(series_list) <= 1:
            return True  # 单系列或无系列，默认去重

        # 获取第一个系列的X轴数据作为基准
        first_x_data = series_list[0].get('x_data', [])

        # 检查其他系列的X轴数据是否与第一个系列相同
        for series in series_list[1:]:
            x_data = series.get('x_data', [])
            if x_data != first_x_data:
                return False  # 发现不同的X轴数据，不应该去重

        return True  # 所有系列的X轴数据都相同，应该去重
    