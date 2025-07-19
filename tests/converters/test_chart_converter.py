
import pytest
from unittest.mock import MagicMock, patch
from src.converters.chart_converter import ChartConverter
from src.models.table_model import Chart, Sheet, ChartPosition

@pytest.fixture
def chart_converter():
    """Fixture for ChartConverter."""
    return ChartConverter(MagicMock())

@pytest.fixture
def sample_chart():
    """Fixture for a sample chart."""
    chart = Chart(name="TestChart", type="bar", anchor="A1")
    chart.chart_data = {"type": "bar", "series": []}
    return chart

@patch('src.converters.chart_converter.SVGChartRenderer')
def test_render_chart_content_success(mock_renderer, chart_converter, sample_chart):
    """Test successful rendering of chart content."""
    mock_renderer.return_value.render_chart_to_svg.return_value = "<svg>chart</svg>"
    html = chart_converter._render_chart_content(sample_chart)
    assert "<svg>chart</svg>" in html

def test_render_chart_content_no_data(chart_converter):
    """Test rendering chart content with no data."""
    chart = Chart(name="NoDataChart", type="bar", anchor="A1")
    html = chart_converter._render_chart_content(chart)
    assert "Chart data not available" in html

@patch('src.converters.chart_converter.create_position_calculator')
def test_generate_overlay_charts_html(mock_pos_calc, chart_converter, sample_chart):
    """Test generating HTML for overlay charts."""
    sample_chart.position = ChartPosition(from_col=0, from_row=0, to_col=5, to_row=10, from_col_offset=0, from_row_offset=0, to_col_offset=0, to_row_offset=0)
    sheet = Sheet(name="Sheet1", rows=[], charts=[sample_chart])
    mock_calculator = MagicMock()
    css_pos = MagicMock()
    css_pos.width = 300
    css_pos.height = 200
    mock_calculator.calculate_chart_css_position.return_value = css_pos
    mock_calculator.generate_chart_html_with_positioning.return_value = "<div>positioned chart</div>"
    mock_pos_calc.return_value = mock_calculator

    html = chart_converter.generate_overlay_charts_html(sheet)
    assert "<div>positioned chart</div>" in html

def test_generate_standalone_charts_html(chart_converter, sample_chart):
    """Test generating HTML for standalone charts."""
    html = chart_converter.generate_standalone_charts_html([sample_chart])
    assert "<h2>Charts</h2>" in html
    assert "<h3>TestChart</h3>" in html
