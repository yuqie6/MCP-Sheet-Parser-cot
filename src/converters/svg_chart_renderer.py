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
from src.utils.color_utils import DEFAULT_CHART_COLORS, normalize_color


class SVGChartRenderer:
    """将Excel图表数据转换为SVG的渲染器。"""
    
    def __init__(self, width: int = 800, height: int = 500, show_axes: bool = False):  # 默认不显示坐标轴
        """
        初始化SVG渲染器。

        参数：
            width: SVG图表宽度（像素）- 基于Excel默认15cm
            height: SVG图表高度（像素）- 基于Excel默认7.5cm
            show_axes: 是否显示坐标轴和网格线（默认False，匹配Excel简洁样式）
        """
        # 基于搜索结果：Excel默认图表大小15x7.5cm
        self.width = width
        self.height = height
        self.show_axes = show_axes
        # 调整边距：如果不显示坐标轴，减少边距
        if show_axes:
            self.margin = {'top': 50, 'right': 80, 'bottom': 60, 'left': 80}
        else:
            self.margin = {'top': 20, 'right': 20, 'bottom': 20, 'left': 20}  # 简洁模式，最小边距
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
        Gets series colors, prioritizing original Excel colors.
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
        """获取图表的CSS样式。"""
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
    
    def _render_bar_chart(self, chart_data: Dict[str, Any]) -> str:
        """Renders a bar chart."""
        svg = self._create_svg_root(chart_data.get('title', ''))
        series_list = chart_data.get('series', [])

        # 保存当前系列，供其他方法使用
        self.current_series = series_list

        # 修复：确保颜色列表与系列列表匹配
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

        unique_x_labels = list(dict.fromkeys(all_x_labels))
        # 修复：Y轴从0开始，这样所有柱子都有合理的高度
        y_min = 0  # 强制从0开始
        y_max = chart_data.get('y_axis_max') if chart_data.get('y_axis_max') is not None else max(all_y_values) if all_y_values else 1
        # 确保y_min和y_max是数值类型
        y_min = float(y_min)
        y_max = float(y_max)
        y_range = y_max - y_min if y_max != y_min else 1

        # 绘制X轴标签（保持Excel样式）
        self._draw_x_axis_labels(svg, unique_x_labels)

        colors = self._get_series_colors(chart_data)
        # 修复：每个数据点都是独立的柱子，不按标签分组
        total_bars = len(all_x_labels)  # 总柱子数
        bar_group_width = self.plot_width / total_bars
        bar_width = bar_group_width * 0.5  # 减小柱子宽度，使其更苗条
        
        # 重新设计：每个数据点都是独立的柱子，按顺序排列
        bar_index = 0  # 全局柱子索引

        for series_idx, series in enumerate(series_list):
            x_data = series.get('x_data', [])
            y_data = series.get('y_data', [])

            for i, (x_label, y_value) in enumerate(zip(x_data, y_data)):
                # 每个柱子都有独立的位置，不按标签分组
                color = normalize_color(series.get('color') or colors[series_idx % len(colors)])

                # 计算柱子位置：按数据顺序排列
                bar_center_x = self.margin['left'] + (bar_index + 0.5) * bar_group_width
                bar_x = bar_center_x - bar_width / 2

                # 计算柱子高度
                y_value_norm = max(0, y_value - y_min)  # 确保非负
                bar_height = (y_value_norm / y_range) * self.plot_height if y_range > 0 else 0
                bar_height = max(5, bar_height)  # 确保至少5像素高度，让柱子可见

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

                # 只有在Excel明确设置显示数据标签时才显示数字
                show_data_labels = (
                    chart_data.get('show_data_labels', False) or
                    series.get('show_data_labels', False) or
                    chart_data.get('data_labels', {}).get('show', False)
                )

                if show_data_labels and bar_height > 8:
                    text = ET.SubElement(svg, 'text', {
                        'x': str(bar_x + bar_width / 2),
                        'y': str(bar_y + 10),
                        'text-anchor': 'middle',
                        'alignment-baseline': 'middle',
                        'class': 'axis-label',
                        'fill': 'white',
                        'font-size': '10px'
                    })
                    text.text = str(int(y_value))

                bar_index += 1  # 下一个柱子
        
        # 只有在Excel明确设置显示图例时才绘制图例
        show_legend = (
            chart_data.get('show_legend', False) or
            chart_data.get('legend', {}).get('show', False)
        )
        if show_legend:
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
                        'r': '3',  # 添加半径属性
                        'fill': color,
                        'class': 'line-point'
                    })
        
        # 只有在Excel明确设置显示图例时才绘制图例
        show_legend = (
            chart_data.get('show_legend', False) or
            chart_data.get('legend', {}).get('show', False)
        )
        if show_legend:
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
            color = normalize_color(colors[i % len(colors)])
            
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
        
        # 只有在Excel明确设置显示图例时才绘制图例
        show_legend = (
            chart_data.get('show_legend', False) or
            chart_data.get('legend', {}).get('show', False)
        )
        if show_legend:
            self._draw_legend(svg, series_list, colors)
        
        return self._format_svg(svg)
    
    def _render_image_chart(self, chart_data: Dict[str, Any]) -> str:
        """渲染图像 - 支持真实图片和占位符。"""
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
    
    def _draw_x_axis_labels(self, svg: ET.Element, x_labels: List[str]):
        """只绘制X轴标签，不绘制坐标轴线和网格线。"""
        # 获取所有标签（包括重复的）
        all_labels = []
        for series in self.current_series:
            all_labels.extend(series.get('x_data', []))

        # 每个柱子都有自己的标签
        bar_group_width = self.plot_width / len(all_labels)

        # X轴标签
        for i, label in enumerate(all_labels):
            # 计算标签位置（居中于每个柱子）
            x_pos = self.margin['left'] + (i + 0.5) * bar_group_width

            text = ET.SubElement(svg, 'text', {
                'x': str(x_pos),
                'y': str(self.margin['top'] + self.plot_height + 15),
                'text-anchor': 'middle',
                'class': 'axis-label',
                'fill': '#666',  # 确保标签可见
                'font-size': '10px'  # 稍微减小字体，避免重叠
            })
            text.text = str(label)

    def _draw_axes(self, svg: ET.Element, x_labels: List[str], y_min: float, y_max: float):
        """绘制完整坐标轴，包括轴线、网格线和标签。"""
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
    
    def _draw_legend(self, svg: ET.Element, series_list: List[Dict], colors: List[str]):
        """绘制图例。"""
        legend_x = self.width - self.margin['right'] + 10
        legend_y = self.margin['top']
        
        for i, series in enumerate(series_list):
            color = normalize_color(colors[i % len(colors)])
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