import logging

from src.converters.svg_chart_renderer import SVGChartRenderer
from src.models.table_model import Chart
from src.utils.chart_positioning import create_position_calculator

logger = logging.getLogger(__name__)


class ChartConverter:
    """Handles chart rendering and positioning."""

    def __init__(self, cell_converter):
        self.cell_converter = cell_converter

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
            dynamic_svg_renderer = SVGChartRenderer(width=int(css_pos.width), height=int(css_pos.height * 1.33))
            chart_content = []
            if chart.chart_data:
                try:
                    svg_content = dynamic_svg_renderer.render_chart_to_svg(chart.chart_data)
                    chart_content.append(svg_content)
                except Exception as e:
                    logger.warning(f"Failed to render overlay chart '{chart.name}' as SVG: {e}")
                    chart_content.append(f'<div class="chart-error">Chart rendering failed</div>')
            else:
                chart_content.append(f'<div class="chart-placeholder">Chart data not available</div>')

            positioned_html = position_calculator.generate_chart_html_with_positioning(chart,
                                                                                       '\n'.join(chart_content))
            charts_html_parts.append(positioned_html)

        return '\n'.join(charts_html_parts)

    def generate_standalone_charts_html(self, charts: list[Chart]) -> str:
        """
        Generates HTML for standalone charts (charts without positioning information).
        """
        standalone_charts = [chart for chart in charts if chart.position is None]
        if not standalone_charts:
            return ""

        charts_html_parts = ['<h2>Charts</h2>']
        svg_renderer = SVGChartRenderer()
        for chart in standalone_charts:
            chart_html = [
                f'<div class="chart-container" id="chart-{chart.name.replace(" ", "-")}" data-anchor="{chart.anchor}">',
                f'<h3>{self.cell_converter._escape_html(chart.name)}</h3>'
            ]
            if chart.chart_data:
                try:
                    svg_content = svg_renderer.render_chart_to_svg(chart.chart_data)
                    chart_html.append(f'<div class="chart-svg-wrapper">{svg_content}</div>')
                except Exception as e:
                    logger.warning(f"Failed to render chart '{chart.name}' as SVG: {e}")
                    chart_html.append(f'<div class="chart-error">Chart rendering failed</div>')
            else:
                chart_html.append(f'<div class="chart-placeholder">Chart data not available</div>')

            chart_html.append('</div>')
            charts_html_parts.extend(chart_html)

        return "\n".join(charts_html_parts)
