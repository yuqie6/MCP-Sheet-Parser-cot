
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

# === TDD测试：提升SVG图表渲染器覆盖率 ===

def test_render_chart_with_size_info_missing_dimensions():
    """
    TDD测试：render_chart_to_svg应该处理size信息不完整的情况

    这个测试覆盖第48-50行的代码路径，当size存在但缺少width_px或height_px时
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer(width=800, height=500)

    # 测试只有width_px的情况
    chart_data = {
        'type': 'bar',
        'title': 'Incomplete Size Chart',
        'size': {
            'width_px': 600
            # 缺少height_px
        },
        'series': [{'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]}]
    }

    svg = renderer.render_chart_to_svg(chart_data)
    # 应该使用默认尺寸（注意实际格式是带px的）
    assert 'width="800px"' in svg
    assert 'height="500px"' in svg

def test_renderer_initialization_with_show_axes_true():
    """
    TDD测试：SVGChartRenderer初始化时show_axes=True应该设置正确的边距

    这个测试覆盖第32-33行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer(width=800, height=500, show_axes=True)

    # 验证边距设置
    assert renderer.margin['top'] == 70
    assert renderer.margin['right'] == 80
    assert renderer.margin['bottom'] == 60
    assert renderer.margin['left'] == 80

    # 验证绘图区域计算
    expected_plot_width = 800 - 80 - 80  # width - left - right
    expected_plot_height = 500 - 70 - 60  # height - top - bottom
    assert renderer.plot_width == expected_plot_width
    assert renderer.plot_height == expected_plot_height

def test_renderer_initialization_with_show_axes_false():
    """
    TDD测试：SVGChartRenderer初始化时show_axes=False应该设置简洁边距

    这个测试覆盖第34-35行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer(width=800, height=500, show_axes=False)

    # 验证简洁模式的边距设置
    assert renderer.margin['top'] == 40
    assert renderer.margin['right'] == 20
    assert renderer.margin['bottom'] == 20
    assert renderer.margin['left'] == 20

    # 验证绘图区域计算
    expected_plot_width = 800 - 20 - 20  # width - left - right
    expected_plot_height = 500 - 40 - 20  # height - top - bottom
    assert renderer.plot_width == expected_plot_width
    assert renderer.plot_height == expected_plot_height

def test_render_legend_if_needed_with_show_legend_true():
    """
    TDD测试：_render_legend_if_needed应该在show_legend为True时渲染图例

    这个测试覆盖第72-94行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    # 创建模拟的SVG元素
    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    chart_data = {
        'show_legend': True,
        'legend': {
            'position': 'right'
        }
    }

    series_list = [
        {'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]},
        {'name': 'Series 2', 'x_data': ['A'], 'y_data': [20]}
    ]

    colors = ['#FF0000', '#00FF00']

    # 模拟_draw_legend方法
    with patch.object(renderer, '_draw_legend') as mock_draw_legend:
        renderer._render_legend_if_needed(svg, chart_data, series_list, colors)

        # 验证_draw_legend被调用
        mock_draw_legend.assert_called_once()

        # 验证调用参数
        call_args = mock_draw_legend.call_args
        assert call_args[0][0] == svg  # svg元素
        assert call_args[0][1] == series_list  # 系列列表
        assert call_args[0][2] == colors  # 颜色列表

def test_render_legend_if_needed_with_legend_enabled():
    """
    TDD测试：_render_legend_if_needed应该在legend.enabled为True时渲染图例

    这个测试覆盖第76行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    chart_data = {
        'legend': {
            'enabled': True,
            'position': 'bottom'
        }
    }

    series_list = [{'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]}]
    colors = ['#FF0000']

    with patch.object(renderer, '_draw_legend') as mock_draw_legend:
        renderer._render_legend_if_needed(svg, chart_data, series_list, colors)
        mock_draw_legend.assert_called_once()

def test_render_legend_if_needed_with_legend_show():
    """
    TDD测试：_render_legend_if_needed应该在legend.show为True时渲染图例

    这个测试覆盖第75行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    chart_data = {
        'legend': {
            'show': True,
            'position': 'top'
        }
    }

    series_list = [{'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]}]
    colors = ['#FF0000']

    with patch.object(renderer, '_draw_legend') as mock_draw_legend:
        renderer._render_legend_if_needed(svg, chart_data, series_list, colors)
        mock_draw_legend.assert_called_once()

def test_render_legend_if_needed_no_legend():
    """
    TDD测试：_render_legend_if_needed应该在没有图例标志时不渲染图例

    这个测试确保方法在没有图例要求时不调用_draw_legend
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    chart_data = {
        'show_legend': False,
        'legend': {
            'enabled': False,
            'show': False
        }
    }

    series_list = [{'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]}]
    colors = ['#FF0000']

    with patch.object(renderer, '_draw_legend') as mock_draw_legend:
        renderer._render_legend_if_needed(svg, chart_data, series_list, colors)

        # 验证_draw_legend没有被调用
        mock_draw_legend.assert_not_called()

# === TDD测试：测试图像渲染功能 ===

def test_render_image_chart_with_nested_image_data():
    """
    TDD测试：_render_image_chart应该能处理嵌套的图像数据结构

    这个测试覆盖第806-812行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    # 测试嵌套字典结构的图像数据
    chart_data = {
        'type': 'image',
        'title': 'Nested Image',
        'image_data': {
            'image_data': b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n\x1a\n\x00\x00\x00\x00IEND\xaeB`\x82'
        }
    }

    result = renderer._render_image_chart(chart_data)

    # 应该返回包含base64编码图像的HTML
    assert '<img src="data:image/png;base64,' in result
    assert 'Nested Image' in result

def test_render_image_chart_with_direct_bytes():
    """
    TDD测试：_render_image_chart应该能处理直接的字节数据

    这个测试覆盖第812行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    # 测试直接的字节数据
    chart_data = {
        'type': 'image',
        'title': 'Direct Bytes',
        'image_data': b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n\x1a\n\x00\x00\x00\x00IEND\xaeB`\x82'
    }

    result = renderer._render_image_chart(chart_data)

    # 应该返回包含base64编码图像的HTML
    assert '<img src="data:image/png;base64,' in result
    assert 'Direct Bytes' in result

def test_render_image_chart_with_jpeg_format():
    """
    TDD测试：_render_image_chart应该能识别JPEG格式

    这个测试覆盖JPEG格式检测的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    # JPEG文件的魔数
    jpeg_data = b'\xff\xd8\xff\xe0' + b'\x00' * 20  # 简化的JPEG数据

    chart_data = {
        'type': 'image',
        'title': 'JPEG Image',
        'image_data': jpeg_data
    }

    result = renderer._render_image_chart(chart_data)

    # 应该识别为JPEG格式
    assert '<img src="data:image/jpeg;base64,' in result

def test_render_image_chart_with_small_data():
    """
    TDD测试：_render_image_chart应该处理太小的图像数据

    这个测试确保方法在图像数据太小时返回占位符
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    chart_data = {
        'type': 'image',
        'title': 'Small Data',
        'image_data': b'small'  # 少于10字节
    }

    result = renderer._render_image_chart(chart_data)

    # 应该返回占位符SVG而不是图像
    assert '<svg' in result
    assert 'Small Data' in result
    assert '<img' not in result

def test_render_image_chart_with_no_image_data():
    """
    TDD测试：_render_image_chart应该处理没有图像数据的情况

    这个测试确保方法在没有图像数据时返回占位符
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    chart_data = {
        'type': 'image',
        'title': 'No Image Data'
        # 没有image_data字段
    }

    result = renderer._render_image_chart(chart_data)

    # 应该返回占位符SVG
    assert '<svg' in result
    assert 'No Image Data' in result
    assert '<img' not in result

def test_render_fallback_chart_with_custom_title():
    """
    TDD测试：_render_fallback_chart应该支持自定义标题

    这个测试覆盖fallback图表的标题处理
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    chart_data = {
        'type': 'unknown_type',
        'title': 'Custom Fallback Title'
    }

    result = renderer._render_fallback_chart(chart_data)

    # 应该包含自定义标题
    assert 'Custom Fallback Title' in result
    assert "Chart type 'unknown_type' not supported" in result

def test_render_fallback_chart_without_title():
    """
    TDD测试：_render_fallback_chart应该处理没有标题的情况

    这个测试确保方法在没有标题时使用默认标题
    """
    # 🔴 红阶段：编写测试描述期望的行为
    renderer = SVGChartRenderer()

    chart_data = {
        'type': 'unknown_type'
        # 没有title字段
    }

    result = renderer._render_fallback_chart(chart_data)

    # 应该使用默认标题
    assert 'Unsupported Chart' in result
    assert "Chart type 'unknown_type' not supported" in result
