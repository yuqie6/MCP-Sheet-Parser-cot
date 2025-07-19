
import pytest
from unittest.mock import MagicMock, patch
from src.converters.svg_chart_renderer import SVGChartRenderer

@pytest.fixture
def renderer():
    """Fixture for SVGChartRenderer."""
    return SVGChartRenderer()

@pytest.fixture
def chart_data():
    """Fixture for sample chart data."""
    return {
        'type': 'bar',
        'title': 'Test Chart',
        'series': [
            {
                'name': 'Series 1',
                'x_data': ['A', 'B', 'C'],
                'y_data': [10, 20, 30]
            }
        ]
    }

def test_render_chart_to_svg_routing(renderer):
    """Test that render_chart_to_svg routes to the correct render method."""
    with patch.object(renderer, '_render_bar_chart') as mock_bar, \
         patch.object(renderer, '_render_line_chart') as mock_line, \
         patch.object(renderer, '_render_pie_chart') as mock_pie, \
         patch.object(renderer, '_render_area_chart') as mock_area, \
         patch.object(renderer, '_render_image_chart') as mock_image, \
         patch.object(renderer, '_render_fallback_chart') as mock_fallback:

        renderer.render_chart_to_svg({'type': 'bar'})
        mock_bar.assert_called_once()

        renderer.render_chart_to_svg({'type': 'line'})
        mock_line.assert_called_once()

        renderer.render_chart_to_svg({'type': 'pie'})
        mock_pie.assert_called_once()

        renderer.render_chart_to_svg({'type': 'area'})
        mock_area.assert_called_once()

        renderer.render_chart_to_svg({'type': 'image'})
        mock_image.assert_called_once()

        renderer.render_chart_to_svg({'type': 'unsupported'})
        mock_fallback.assert_called_once()

def test_render_bar_chart(renderer, chart_data):
    """Test rendering a bar chart."""
    svg = renderer.render_chart_to_svg(chart_data)
    assert '<svg' in svg
    assert 'class="bar-rect"' in svg
    assert 'Test Chart' in svg

def test_render_line_chart(renderer, chart_data):
    """Test rendering a line chart."""
    chart_data['type'] = 'line'
    renderer.current_series = chart_data['series']
    svg = renderer.render_chart_to_svg(chart_data)
    assert '<svg' in svg
    assert 'class="line-path"' in svg

def test_render_pie_chart(renderer, chart_data):
    """Test rendering a pie chart."""
    chart_data['type'] = 'pie'
    svg = renderer.render_chart_to_svg(chart_data)
    assert '<svg' in svg
    assert 'class="pie-slice"' in svg

def test_create_svg_root_with_title(renderer):
    """Test _create_svg_root with a title."""
    svg_root = renderer._create_svg_root("My Title")
    assert svg_root.tag == 'svg'
    title_elem = svg_root.find('.//text[@class="chart-title"]')
    assert title_elem is not None
    assert title_elem.text == "My Title"

def test_get_series_colors(renderer, chart_data):
    """Test _get_series_colors."""
    # Default colors
    colors = renderer._get_series_colors(chart_data)
    assert isinstance(colors, list)
    # Custom colors
    chart_data['colors'] = ['#FF0000', '#00FF00']
    colors = renderer._get_series_colors(chart_data)
    assert colors == ['#FF0000', '#00FF00']

def test_should_deduplicate_x_labels(renderer):
    """Test _should_deduplicate_x_labels."""
    series_same_x = [
        {'x_data': ['A', 'B']},
        {'x_data': ['A', 'B']}
    ]
    series_diff_x = [
        {'x_data': ['A', 'B']},
        {'x_data': ['C', 'D']}
    ]
    assert renderer._should_deduplicate_x_labels(series_same_x) is True
    assert renderer._should_deduplicate_x_labels(series_diff_x) is False

def test_render_area_chart(renderer, chart_data):
    """Test rendering an area chart."""
    chart_data['type'] = 'area'
    renderer.current_series = chart_data['series']
    svg = renderer.render_chart_to_svg(chart_data)
    assert '<svg' in svg
    assert 'class="area-path"' in svg

def test_render_image_chart(renderer):
    """Test rendering an image chart."""
    chart_data = {'type': 'image', 'image_data': b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n\x1a\n\x00\x00\x00\x00IEND\xaeB`\x82'}
    html = renderer.render_chart_to_svg(chart_data)
    assert '<img src="data:image/png;base64,' in html

def test_render_fallback_chart(renderer):
    """Test rendering a fallback chart for unsupported types."""
    chart_data = {'type': 'radar', 'title': 'Unsupported Chart'}
    svg = renderer.render_chart_to_svg(chart_data)
    assert "Chart type 'radar' not supported" in svg
