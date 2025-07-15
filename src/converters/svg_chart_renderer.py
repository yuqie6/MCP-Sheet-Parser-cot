"""
SVG图表渲染器 - 将Excel图表转换为高质量的SVG格式。

这个模块替代了原有的matplotlib PNG渲染方案，提供：
1. 矢量图形质量
2. 更好的位置控制
3. CSS样式支持
4. 响应式设计
"""

from typing import Any, Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET
from xml.dom import minidom


# Excel主题颜色映射（Office 2016+ 默认主题）
EXCEL_THEME_COLORS = {
    'accent1': '#5B9BD5',  # 蓝色
    'accent2': '#70AD47',  # 绿色  
    'accent3': '#FFC000',  # 橙色
    'accent4': '#E15759',  # 红色
    'accent5': '#4472C4',  # 深蓝色
    'accent6': '#FF6B35',  # 橙红色
    'dk1': '#000000',      # 深色1（黑色）
    'lt1': '#FFFFFF',      # 浅色1（白色）
    'dk2': '#44546A',      # 深色2（深灰蓝）
    'lt2': '#E7E6E6',      # 浅色2（浅灰）
    'bg1': '#FFFFFF',      # 背景1
    'bg2': '#E7E6E6',      # 背景2
    'tx1': '#000000',      # 文本1
    'tx2': '#44546A',      # 文本2
}

# 图表系列的默认颜色顺序（基于Excel主题）
DEFAULT_CHART_COLORS = [
    EXCEL_THEME_COLORS['accent1'],  # 蓝色
    EXCEL_THEME_COLORS['accent2'],  # 绿色
    EXCEL_THEME_COLORS['accent3'],  # 橙色
    EXCEL_THEME_COLORS['accent4'],  # 红色
    EXCEL_THEME_COLORS['accent5'],  # 深蓝色
    EXCEL_THEME_COLORS['accent6'],  # 橙红色
]


class SVGChartRenderer:
    """将Excel图表数据转换为SVG的渲染器。"""
    
    def __init__(self, width: int = 400, height: int = 300):
        """
        初始化SVG渲染器。
        
        参数：
            width: SVG图表宽度（像素）
            height: SVG图表高度（像素）
        """
        self.width = width
        self.height = height
        self.margin = {'top': 40, 'right': 40, 'bottom': 60, 'left': 80}
        self.plot_width = width - self.margin['left'] - self.margin['right']
        self.plot_height = height - self.margin['top'] - self.margin['bottom']
        
    def render_chart_to_svg(self, chart_data: Dict[str, Any]) -> str:
        """
        将图表数据渲染为SVG字符串。
        
        参数：
            chart_data: 包含图表类型、数据、标题等信息的字典
                {
                    'type': 'bar'|'line'|'pie'|'area',
                    'title': '图表标题',
                    'series': [{'name': '系列名', 'x_data': [...], 'y_data': [...], 'color': '#RRGGBB'}],
                    'x_axis_title': 'X轴标题',
                    'y_axis_title': 'Y轴标题',
                    'colors': ['#RRGGBB', ...],  # 原始Excel颜色
                    'size': {'width_px': 400, 'height_px': 300}  # 原始Excel尺寸
                }
        
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
    
    def _create_svg_root(self, title: str = "") -> ET.Element:
        """创建SVG根元素。"""
        svg = ET.Element('svg', {
            'width': str(self.width),
            'height': str(self.height),
            'viewBox': f'0 0 {self.width} {self.height}',
            'xmlns': 'http://www.w3.org/2000/svg',
            'class': 'excel-chart-svg'
        })
        
        # 添加样式定义
        style = ET.SubElement(svg, 'style')
        style.text = self._get_chart_css()
        
        # 添加标题
        if title:
            title_elem = ET.SubElement(svg, 'text', {
                'x': str(self.width // 2),
                'y': '25',
                'text-anchor': 'middle',
                'class': 'chart-title'
            })
            title_elem.text = title
        
        return svg
    
    def _get_series_colors(self, chart_data: Dict[str, Any]) -> List[str]:
        """
        获取系列颜色，优先使用Excel原始颜色。
        
        参数：
            chart_data: 图表数据
            
        返回：
            颜色列表
        """
        # 优先使用从Excel提取的颜色
        if 'colors' in chart_data and chart_data['colors']:
            return chart_data['colors']
        
        # 从系列数据中提取颜色
        colors = []
        for series in chart_data.get('series', []):
            if 'color' in series and series['color']:
                colors.append(series['color'])
        
        if colors:
            return colors
        
        # 回退到默认颜色
        return DEFAULT_CHART_COLORS
    
    def _get_chart_css(self) -> str:
        """获取图表的CSS样式。"""
        return """
        .excel-chart-svg {
            font-family: 'Microsoft YaHei', 'SimHei', 'PingFang SC', 'Hiragino Sans GB', 'Source Han Sans SC', 'Noto Sans CJK SC', 'Segoe UI', Arial, sans-serif;
        }
        .chart-title {
            font-size: 16px;
            font-weight: bold;
            fill: #333;
        }
        .axis-title {
            font-size: 12px;
            fill: #666;
        }
        .axis-label {
            font-size: 10px;
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
            font-size: 10px;
            fill: #666;
        }
        """
    
    def _render_bar_chart(self, chart_data: Dict[str, Any]) -> str:
        """渲染柱状图。"""
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
        
        # 去重并排序x标签
        unique_x_labels = list(dict.fromkeys(all_x_labels))  # 保持顺序的去重
        
        # 计算y轴范围，确保从0开始显示
        y_min = 0  # 强制从0开始，这样所有数值都能正确显示高度
        y_max = max(all_y_values) if all_y_values else 1
        y_range = y_max - y_min if y_max != y_min else 1
        
        # 绘制坐标轴
        self._draw_axes(svg, unique_x_labels, y_min, y_max)
        
        # 绘制数据系列
        colors = self._get_series_colors(chart_data)
        # 修复：正确计算柱子宽度
        total_bar_space = self.plot_width * 0.8  # 怰80%的空间用于柱子
        bar_group_width = total_bar_space / len(unique_x_labels)
        single_bar_width = bar_group_width / len(series_list) * 0.8  # 柱子之间留空隙
        
        for series_idx, series in enumerate(series_list):
            color = colors[series_idx % len(colors)]
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])
            
            for i, (x_label, y_value) in enumerate(zip(x_data, y_data)):
                if x_label in unique_x_labels:
                    x_pos = unique_x_labels.index(x_label)
                    # 修复：正确计算柱子位置
                    group_start = self.margin['left'] + x_pos * (self.plot_width / len(unique_x_labels))
                    group_center = group_start + (self.plot_width / len(unique_x_labels)) / 2
                    total_series_width = single_bar_width * len(series_list)
                    bar_x = group_center - total_series_width / 2 + series_idx * single_bar_width
                    
                    # 修复：确保非零值有可见的柱子高度
                    if y_value == 0:
                        bar_height = 0
                    else:
                        # 计算柱子高度，确保最小可见高度
                        calculated_height = (y_value - y_min) / y_range * self.plot_height
                        bar_height = max(calculated_height, 3)  # 最小3像素高度
                    
                    bar_y = self.margin['top'] + self.plot_height - bar_height
                    
                    # 创建柱子
                    rect = ET.SubElement(svg, 'rect', {
                        'x': str(bar_x),
                        'y': str(bar_y),
                        'width': str(single_bar_width),
                        'height': str(bar_height),
                        'fill': color,
                        'class': 'bar-rect'
                    })
                    
                    # 添加数值标签
                    if bar_height > 15:  # 只在足够高的柱子上显示标签
                        text = ET.SubElement(svg, 'text', {
                            'x': str(bar_x + single_bar_width / 2),
                            'y': str(bar_y + bar_height / 2),
                            'text-anchor': 'middle',
                            'alignment-baseline': 'middle',
                            'class': 'axis-label',
                            'fill': 'white'
                        })
                        text.text = str(y_value)
        
        # 绘制图例
        self._draw_legend(svg, series_list, colors)
        
        return self._format_svg(svg)
    
    def _render_line_chart(self, chart_data: Dict[str, Any]) -> str:
        """渲染折线图。"""
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
        
        for series_idx, series in enumerate(series_list):
            color = colors[series_idx % len(colors)]
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])
            
            if len(x_data) < 2 or len(y_data) < 2:
                continue
            
            # 创建路径点
            points = []
            for x_label, y_value in zip(x_data, y_data):
                if x_label in unique_x_labels:
                    x_pos = unique_x_labels.index(x_label)
                    # 修复：正确计算点的x坐标
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
                        'fill': color,
                        'class': 'line-point'
                    })
        
        # 绘制图例
        self._draw_legend(svg, series_list, colors)
        
        return self._format_svg(svg)
    
    def _render_pie_chart(self, chart_data: Dict[str, Any]) -> str:
        """渲染饼图。"""
        svg = self._create_svg_root(chart_data.get('title', ''))
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
        
        center_x = self.width // 2
        center_y = self.height // 2
        radius = min(self.plot_width, self.plot_height) // 2 - 20
        
        # 优先使用从Excel提取的颜色
        colors = []
        if series.get('colors'):  # 饼图的每个片段可能有不同颜色
            colors = series['colors']
        elif 'colors' in chart_data and chart_data['colors']:
            colors = chart_data['colors']
        else:
            colors = DEFAULT_CHART_COLORS + ['#A5A5A5', '#70E000']  # 扩展颜色用于多系列饼图
        
        current_angle = 0
        for i, (label, value) in enumerate(zip(labels, values)):
            angle = (value / total) * 360
            color = colors[i % len(colors)]
            
            # 计算扇形路径
            start_angle_rad = current_angle * 3.14159 / 180
            end_angle_rad = (current_angle + angle) * 3.14159 / 180
            
            start_x = center_x + radius * cos(start_angle_rad)
            start_y = center_y + radius * sin(start_angle_rad)
            end_x = center_x + radius * cos(end_angle_rad)
            end_y = center_y + radius * sin(end_angle_rad)
            
            large_arc = 1 if angle > 180 else 0
            
            path_data = f"M {center_x} {center_y} L {start_x} {start_y} A {radius} {radius} 0 {large_arc} 1 {end_x} {end_y} Z"
            
            path = ET.SubElement(svg, 'path', {
                'd': path_data,
                'fill': color,
                'class': 'pie-slice'
            })
            
            # 添加标签
            mid_angle = (current_angle + angle / 2) * 3.14159 / 180
            label_x = center_x + (radius + 20) * cos(mid_angle)
            label_y = center_y + (radius + 20) * sin(mid_angle)
            
            text = ET.SubElement(svg, 'text', {
                'x': str(label_x),
                'y': str(label_y),
                'text-anchor': 'middle',
                'class': 'axis-label'
            })
            percentage = (value / total) * 100
            text.text = f"{label} ({percentage:.1f}%)"
            
            current_angle += angle
        
        return self._format_svg(svg)
    
    def _render_area_chart(self, chart_data: Dict[str, Any]) -> str:
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
            color = colors[series_idx % len(colors)]
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
        
        # 绘制图例
        self._draw_legend(svg, series_list, colors)
        
        return self._format_svg(svg)
    
    def _render_image_chart(self, chart_data: Dict[str, Any]) -> str:
        """渲染图像占位符。"""
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
        if chart_data.get('image_data'):
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
    
    def _render_fallback_chart(self, chart_data: Dict[str, Any]) -> str:
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
    
    def _draw_axes(self, svg: ET.Element, x_labels: List[str], y_min: float, y_max: float):
        """绘制坐标轴。"""
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
        
        # X轴标签
        for i, label in enumerate(x_labels):
            # 修复：正确计算X轴标签位置
            if len(x_labels) == 1:
                x_pos = self.margin['left'] + self.plot_width / 2
            else:
                x_pos = self.margin['left'] + i * (self.plot_width / (len(x_labels) - 1))
            
            text = ET.SubElement(svg, 'text', {
                'x': str(x_pos),
                'y': str(self.margin['top'] + self.plot_height + 15),
                'text-anchor': 'middle',
                'class': 'axis-label'
            })
            text.text = str(label)
        
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
    
    def _draw_legend(self, svg: ET.Element, series_list: List[Dict], colors: List[str]):
        """绘制图例。"""
        legend_x = self.width - self.margin['right'] + 10
        legend_y = self.margin['top']
        
        for i, series in enumerate(series_list):
            color = colors[i % len(colors)]
            series_name = series.get('name', f'Series {i + 1}')
            
            # 图例色块
            rect = ET.SubElement(svg, 'rect', {
                'x': str(legend_x),
                'y': str(legend_y + i * 20),
                'width': '12',
                'height': '12',
                'fill': color
            })
            
            # 图例文字
            text = ET.SubElement(svg, 'text', {
                'x': str(legend_x + 16),
                'y': str(legend_y + i * 20 + 9),
                'class': 'legend-item'
            })
            text.text = series_name
    
    def _format_svg(self, svg: ET.Element) -> str:
        """格式化SVG为字符串。"""
        rough_string = ET.tostring(svg, encoding='unicode')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")[22:]  # 去掉XML声明


def cos(angle_rad: float) -> float:
    """计算余弦值。"""
    import math
    return math.cos(angle_rad)


def sin(angle_rad: float) -> float:
    """计算正弦值。"""
    import math
    return math.sin(angle_rad)