import logging

from src.converters.svg_chart_renderer import SVGChartRenderer
from src.models.table_model import Chart
from src.utils.chart_positioning import create_position_calculator
from src.utils.html_utils import escape_html, create_html_element

logger = logging.getLogger(__name__)


class ChartConverter:
    """Handles chart rendering and positioning."""

    def __init__(self, cell_converter):
        self.cell_converter = cell_converter

    def _render_chart_content(self, chart: Chart, width: int = None, height: int = None) -> str:
        """通用的图表内容渲染方法，减少重复代码。"""
        # 动态确定图表尺寸
        if width is None:
            width = 800
        if height is None:
            height = 500
            
        renderer = SVGChartRenderer(width=width, height=height)
        
        if chart.chart_data:
            try:
                return renderer.render_chart_to_svg(chart.chart_data)
            except Exception as e:
                logger.warning(f"Failed to render chart '{chart.name}' as SVG: {e}")
                return create_html_element('div', f'Chart rendering failed: {str(e)}', css_classes=['chart-error'])
        else:
            return create_html_element('div', 'Chart data not available', css_classes=['chart-placeholder'])

    def generate_overlay_charts_html(self, sheet) -> str:
        """
        Generates HTML for overlay charts (charts with positioning information).
        """
        position_calculator = create_position_calculator(sheet)
        overlay_charts = [chart for chart in sheet.charts if chart.position is not None]
        if not overlay_charts:
            return ""

        charts_html_parts = []
        for chart in overlay_charts:
            css_pos = position_calculator.calculate_chart_css_position(chart.position)
            # 修复：使用与定位器一致的高度计算，确保SVG与容器匹配
            chart_width = max(200, int(css_pos.width))  # width已经是px
            # 使用与chart_positioning.py一致的高度计算
            if chart.type == 'image':
                chart_height = max(150, int(css_pos.height * 1.333))  # 图片高度
            else:
                chart_height = max(150, int(css_pos.height * 1.333))  # 其他图表高度，与容器一致
            
            # 使用新的通用方法渲染图表
            chart_content = self._render_chart_content(chart, chart_width, chart_height)
            positioned_html = position_calculator.generate_chart_html_with_positioning(chart, chart_content)
            charts_html_parts.append(positioned_html)

        return '\n'.join(charts_html_parts)

    def generate_standalone_charts_html(self, charts: list[Chart]) -> str:
        """
        Generates HTML for standalone charts (charts without positioning information).
        """
        standalone_charts = [chart for chart in charts if chart.position is None]
        if not standalone_charts:
            return ""

        charts_html_parts = [create_html_element('h2', 'Charts')]
        for chart in standalone_charts:
            # 创建容器属性
            container_attrs = {
                'id': f'chart-{chart.name.replace(" ", "-")}',
                'data-anchor': chart.anchor
            }
            
            # 创建图表标题
            chart_title = create_html_element('h3', escape_html(chart.name))
            
            # 使用新的通用方法渲染图表
            chart_content = self._render_chart_content(chart)
            chart_wrapper = create_html_element('div', chart_content, css_classes=['chart-svg-wrapper'])
            
            # 组合完整的图表HTML
            chart_container_content = chart_title + chart_wrapper
            chart_container = create_html_element('div', chart_container_content, 
                                                attributes=container_attrs, 
                                                css_classes=['chart-container'])
            charts_html_parts.append(chart_container)

        return "\n".join(charts_html_parts)
