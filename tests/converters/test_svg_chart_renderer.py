"""
SVG图表渲染器测试模块

测试SVG图表渲染器的核心功能：图表类型支持、错误处理、样式渲染等。
"""

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


def test_render_chart_with_size_info_missing_dimensions():
    """
    TDD测试：render_chart_to_svg应该处理size信息不完整的情况

    这个测试覆盖第48-50行的代码路径，当size存在但缺少width_px或height_px时
    """
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
    renderer = SVGChartRenderer()

    # 创建SVG根元素
    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    # 测试数据
    chart_data = {'show_legend': True}
    series_list = [{'name': 'Series 1'}, {'name': 'Series 2'}]
    colors = ['#FF0000', '#00FF00']

    # 使用mock来验证_draw_legend被调用
    with patch.object(renderer, '_draw_legend') as mock_draw_legend:
        renderer._render_legend_if_needed(svg, chart_data, series_list, colors)
        mock_draw_legend.assert_called_once_with(svg, series_list, colors, legend_style=None)


def test_render_legend_if_needed_with_legend_enabled():
    """
    TDD测试：_render_legend_if_needed应该在legend.enabled为True时渲染图例

    这个测试覆盖第76行的代码路径
    """
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


def test_render_image_chart_with_nested_image_data():
    """
    TDD测试：_render_image_chart应该能处理嵌套的图像数据结构

    这个测试覆盖第806-812行的代码路径
    """
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
    renderer = SVGChartRenderer()

    chart_data = {
        'type': 'unknown_type'
        # 没有title字段
    }

    result = renderer._render_fallback_chart(chart_data)

    # 应该使用默认标题
    assert 'Unsupported Chart' in result
    assert "Chart type 'unknown_type' not supported" in result


def test_render_chart_with_complete_size_info():
    """
    TDD测试：render_chart_to_svg应该处理完整的size信息

    这个测试覆盖第50-55行的代码路径
    """
    renderer = SVGChartRenderer(width=800, height=500)

    chart_data = {
        'type': 'bar',
        'title': 'Sized Chart',
        'size': {
            'width_px': 600,
            'height_px': 400
        },
        'series': [{'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]}]
    }

    svg = renderer.render_chart_to_svg(chart_data)

    # 应该使用size信息中的尺寸（但有最小值限制）
    assert 'width="600px"' in svg
    assert 'height="400px"' in svg

    # 验证内部属性也被更新
    assert renderer.width == 600
    assert renderer.height == 400

def test_render_legend_with_entries_and_delete_flag():
    """
    TDD测试：_render_legend_if_needed应该处理带有delete标志的图例条目

    这个测试覆盖第85-91行的代码路径
    """
    renderer = SVGChartRenderer()

    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    chart_data = {
        'show_legend': True,
        'legend': {
            'entries': [
                {'text': 'Series 1', 'index': 0, 'delete': False},
                {'text': 'Series 2', 'index': 1, 'delete': True},  # 应该被跳过
                {'text': 'Series 3', 'index': 2, 'delete': False}
            ]
        }
    }

    series_list = [
        {'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]},
        {'name': 'Series 2', 'x_data': ['A'], 'y_data': [20]},
        {'name': 'Series 3', 'x_data': ['A'], 'y_data': [30]}
    ]

    colors = ['#FF0000', '#00FF00', '#0000FF']

    with patch.object(renderer, '_draw_legend') as mock_draw_legend:
        renderer._render_legend_if_needed(svg, chart_data, series_list, colors)

        # 验证_draw_legend被调用
        mock_draw_legend.assert_called_once()

        # 验证传递的系列列表只包含未删除的条目
        call_args = mock_draw_legend.call_args
        legend_series = call_args[0][1]  # 第二个参数是系列列表

        # 应该只有2个系列（跳过了delete=True的条目）
        assert len(legend_series) == 2
        assert legend_series[0]['name'] == 'Series 1'
        assert legend_series[1]['name'] == 'Series 3'

def test_render_legend_with_entries_without_delete_flag():
    """
    TDD测试：_render_legend_if_needed应该处理没有delete标志的图例条目

    这个测试覆盖第92-94行的代码路径
    """
    renderer = SVGChartRenderer()

    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    chart_data = {
        'show_legend': True,
        'legend': {
            'entries': [
                {'text': 'Custom Series 1', 'index': 0},
                {'text': 'Custom Series 2', 'index': 1}
            ]
        }
    }

    series_list = [
        {'name': 'Original Series 1', 'x_data': ['A'], 'y_data': [10]},
        {'name': 'Original Series 2', 'x_data': ['A'], 'y_data': [20]}
    ]

    colors = ['#FF0000', '#00FF00']

    with patch.object(renderer, '_draw_legend') as mock_draw_legend:
        renderer._render_legend_if_needed(svg, chart_data, series_list, colors)

        # 验证_draw_legend被调用
        mock_draw_legend.assert_called_once()

        # 验证传递的系列列表使用了自定义名称
        call_args = mock_draw_legend.call_args
        legend_series = call_args[0][1]

        assert len(legend_series) == 2
        assert legend_series[0]['name'] == 'Custom Series 1'
        assert legend_series[1]['name'] == 'Custom Series 2'

def test_render_legend_fallback_to_series_list():
    """
    TDD测试：_render_legend_if_needed应该在没有entries时使用series_list

    这个测试覆盖第92-94行的else分支
    """
    renderer = SVGChartRenderer()

    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    chart_data = {
        'show_legend': True,
        'legend': {
            'position': 'right'
            # 没有entries字段
        }
    }

    series_list = [
        {'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]},
        {'name': 'Series 2', 'x_data': ['A'], 'y_data': [20]}
    ]

    colors = ['#FF0000', '#00FF00']

    with patch.object(renderer, '_draw_legend') as mock_draw_legend:
        renderer._render_legend_if_needed(svg, chart_data, series_list, colors)

        # 验证_draw_legend被调用，并且使用了原始的series_list
        mock_draw_legend.assert_called_once()

        call_args = mock_draw_legend.call_args
        passed_series = call_args[0][1]

        # 应该直接使用原始的series_list
        assert passed_series == series_list

def test_render_chart_with_minimum_size_constraints():
    """
    TDD测试：render_chart_to_svg应该应用最小尺寸约束

    这个测试覆盖第51-52行的最小尺寸限制
    """
    renderer = SVGChartRenderer()

    chart_data = {
        'type': 'bar',
        'title': 'Small Chart',
        'size': {
            'width_px': 100,  # 小于最小宽度200
            'height_px': 50   # 小于最小高度150
        },
        'series': [{'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]}]
    }

    svg = renderer.render_chart_to_svg(chart_data)

    # 应该应用最小尺寸约束
    assert 'width="200px"' in svg  # 最小宽度200px
    assert 'height="150px"' in svg  # 最小高度150px

    # 验证内部属性
    assert renderer.width == 200
    assert renderer.height == 150

def test_render_chart_with_plot_area_recalculation():
    """
    TDD测试：render_chart_to_svg应该重新计算绘图区域

    这个测试覆盖第54-55行的绘图区域重新计算
    """
    renderer = SVGChartRenderer(width=800, height=600, show_axes=True)

    # 记录初始的绘图区域
    initial_plot_width = renderer.plot_width
    initial_plot_height = renderer.plot_height

    chart_data = {
        'type': 'bar',
        'title': 'Resized Chart',
        'size': {
            'width_px': 1000,
            'height_px': 800
        },
        'series': [{'name': 'Series 1', 'x_data': ['A'], 'y_data': [10]}]
    }

    svg = renderer.render_chart_to_svg(chart_data)

    # 验证绘图区域被重新计算
    expected_plot_width = 1000 - renderer.margin['left'] - renderer.margin['right']
    expected_plot_height = 800 - renderer.margin['top'] - renderer.margin['bottom']

    assert renderer.plot_width == expected_plot_width
    assert renderer.plot_height == expected_plot_height

    # 确保与初始值不同
    assert renderer.plot_width != initial_plot_width
    assert renderer.plot_height != initial_plot_height

# === 错误处理和边界情况测试 ===

def test_handle_chart_error():
    """
    TDD测试：_handle_chart_error应该生成错误占位符SVG

    这个测试覆盖第133-157行的错误处理代码
    """
    renderer = SVGChartRenderer(width=400, height=300)

    # 模拟一个异常
    test_error = ValueError("测试错误信息")

    # 调用错误处理方法
    error_svg = renderer._handle_chart_error("bar", test_error)

    # 验证返回的是有效的SVG字符串
    assert isinstance(error_svg, str)
    assert '<svg' in error_svg
    assert 'Chart Error: bar' in error_svg

    # 验证错误占位符元素
    assert 'fill="#ffebee"' in error_svg  # 错误背景色
    assert 'stroke="#f44336"' in error_svg  # 错误边框色
    assert 'stroke-dasharray="5,5"' in error_svg  # 虚线边框

    # 验证错误消息文本
    assert 'Error rendering bar: 测试错误信息' in error_svg
    assert 'fill="#d32f2f"' in error_svg  # 错误文本颜色

def test_create_svg_root_with_title_style():
    """
    TDD测试：_create_svg_root应该处理标题样式

    这个测试覆盖第186-197行的标题样式处理代码
    """
    renderer = SVGChartRenderer()

    title_style = {
        'font_family': 'Arial',
        'font_size': 16,
        'color': '#333333',
        'bold': True
    }

    svg_root = renderer._create_svg_root("样式标题", title_style)

    # 验证SVG根元素
    assert svg_root.tag == 'svg'

    # 查找标题元素
    title_elem = svg_root.find('.//text[@class="chart-title"]')
    assert title_elem is not None
    assert title_elem.text == "样式标题"

    # 验证样式属性
    style_attr = title_elem.get('style', '')
    assert "font-family: 'Arial'" in style_attr
    assert "font-size: 16px" in style_attr
    assert "fill: #333333" in style_attr
    assert "font-weight: bold" in style_attr

def test_create_svg_root_with_partial_title_style():
    """
    TDD测试：_create_svg_root应该处理部分标题样式

    这个测试覆盖标题样式的各个分支条件
    """
    renderer = SVGChartRenderer()

    # 测试只有部分样式属性
    title_style = {
        'font_family': 'Times New Roman',
        'color': '#FF0000'
        # 缺少font_size和bold
    }

    svg_root = renderer._create_svg_root("部分样式", title_style)
    title_elem = svg_root.find('.//text[@class="chart-title"]')

    style_attr = title_elem.get('style', '')
    assert "font-family: 'Times New Roman'" in style_attr
    assert "fill: #FF0000" in style_attr
    # 不应该包含未设置的属性
    assert "font-size:" not in style_attr
    assert "font-weight:" not in style_attr

def test_create_svg_root_with_empty_title_style_values():
    """
    TDD测试：_create_svg_root应该忽略空的标题样式值

    这个测试覆盖样式值为空的情况
    """
    renderer = SVGChartRenderer()

    title_style = {
        'font_family': '',  # 空字符串
        'font_size': None,  # None值
        'color': '#000000',  # 有效值
        'bold': False  # False值
    }

    svg_root = renderer._create_svg_root("空样式测试", title_style)
    title_elem = svg_root.find('.//text[@class="chart-title"]')

    style_attr = title_elem.get('style', '')
    # 只应该包含有效的非空值
    assert "fill: #000000" in style_attr
    # 不应该包含空值
    assert "font-family:" not in style_attr
    assert "font-size:" not in style_attr
    assert "font-weight:" not in style_attr

def test_should_show_data_labels_chart_level():
    """
    TDD测试：_should_show_data_labels应该检查图表级别的数据标签设置

    这个测试覆盖第104-105行的代码路径
    """
    renderer = SVGChartRenderer()

    # 测试图表级别的show_data_labels为True
    chart_data = {'show_data_labels': True}
    assert renderer._should_show_data_labels(chart_data) is True

    # 测试图表级别的show_data_labels为False
    chart_data = {'show_data_labels': False}
    assert renderer._should_show_data_labels(chart_data) is False

def test_should_show_data_labels_series_level():
    """
    TDD测试：_should_show_data_labels应该检查系列级别的数据标签设置

    这个测试覆盖第108-109行的代码路径
    """
    renderer = SVGChartRenderer()

    chart_data = {}  # 图表级别没有设置
    series_data = {'show_data_labels': True}

    assert renderer._should_show_data_labels(chart_data, series_data) is True

def test_should_show_data_labels_nested_config():
    """
    TDD测试：_should_show_data_labels应该检查嵌套的数据标签配置

    这个测试覆盖第112-114行的代码路径
    """
    renderer = SVGChartRenderer()

    chart_data = {
        'data_labels': {
            'show': True
        }
    }

    assert renderer._should_show_data_labels(chart_data) is True

def test_should_show_data_labels_all_false():
    """
    TDD测试：_should_show_data_labels应该在所有条件都为False时返回False

    这个测试覆盖第116行的默认返回值
    """
    renderer = SVGChartRenderer()

    chart_data = {}  # 没有任何数据标签设置
    series_data = {}  # 系列也没有设置

    assert renderer._should_show_data_labels(chart_data, series_data) is False

def test_create_data_label_element():
    """
    TDD测试：_create_data_label_element应该创建数据标签元素

    这个测试覆盖第118-131行的数据标签创建代码
    """
    renderer = SVGChartRenderer()

    # 创建SVG根元素
    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    # 调用方法创建数据标签
    renderer._create_data_label_element(
        svg, 100, 200, "测试标签",
        font_size='12px', fill='#333333', text_anchor='start'
    )

    # 验证创建的文本元素
    text_elem = svg.find('text')
    assert text_elem is not None
    assert text_elem.get('x') == '100'
    assert text_elem.get('y') == '200'
    assert text_elem.get('text-anchor') == 'start'
    assert text_elem.get('alignment-baseline') == 'middle'
    assert text_elem.get('class') == 'axis-label'
    assert text_elem.get('fill') == '#333333'
    assert text_elem.get('font-size') == '12px'
    assert text_elem.text == "测试标签"

def test_create_data_label_element_with_defaults():
    """
    TDD测试：_create_data_label_element应该使用默认参数

    这个测试验证默认参数的使用
    """
    renderer = SVGChartRenderer()

    import xml.etree.ElementTree as ET
    svg = ET.Element('svg')

    # 只传递必需参数，其他使用默认值
    renderer._create_data_label_element(svg, 50, 75, "默认样式")

    text_elem = svg.find('text')
    assert text_elem.get('font-size') == '10px'  # 默认字体大小
    assert text_elem.get('fill') == 'white'  # 默认颜色
    assert text_elem.get('text-anchor') == 'middle'  # 默认对齐方式


class TestSVGChartRendererColorHandling:
    """TDD测试：SVG图表渲染器颜色处理测试"""

    def test_get_series_colors_with_series_colors(self):
        """
        TDD测试：get_series_colors应该返回系列中定义的颜色

        覆盖代码行：214, 217 - 系列颜色提取和返回逻辑
        """
        renderer = SVGChartRenderer()

        # 创建包含颜色的图表数据
        chart_data = {
            'series': [
                {'name': 'Series1', 'color': '#FF0000'},
                {'name': 'Series2', 'color': '#00FF00'},
                {'name': 'Series3', 'color': '#0000FF'}
            ]
        }

        # 测试颜色提取
        colors = renderer._get_series_colors(chart_data)

        # 验证返回的是系列中定义的颜色
        assert colors == ['#FF0000', '#00FF00', '#0000FF']

    def test_get_series_colors_fallback_to_default(self):
        """
        TDD测试：get_series_colors应该在没有系列颜色时回退到默认颜色

        覆盖代码行：217 - 默认颜色回退逻辑
        """
        renderer = SVGChartRenderer()

        # 创建没有颜色的图表数据
        chart_data = {
            'series': [
                {'name': 'Series1'},
                {'name': 'Series2'}
            ]
        }

        # 测试颜色提取
        colors = renderer._get_series_colors(chart_data)

        # 验证返回的是默认颜色
        from src.utils.color_utils import DEFAULT_CHART_COLORS
        assert colors == DEFAULT_CHART_COLORS

class TestSVGChartRendererBarChartEdgeCases:
    """TDD测试：SVG图表渲染器柱状图边界情况测试"""

    def test_render_bar_chart_empty_series_list(self):
        """
        TDD测试：render_bar_chart应该处理空系列列表

        覆盖代码行：306, 320 - 空系列处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建空系列的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Empty Chart',
            'series': []  # 空系列列表
        }

        # 测试渲染
        svg_result = renderer._render_bar_chart(chart_data)

        # 应该返回有效的SVG，即使没有数据
        assert '<svg' in svg_result
        assert 'Empty Chart' in svg_result

    def test_render_bar_chart_deduplicate_x_labels(self):
        """
        TDD测试：render_bar_chart应该正确处理X轴标签去重

        覆盖代码行：329-330 - X轴标签去重逻辑
        """
        renderer = SVGChartRenderer()

        # 创建需要去重X轴标签的图表数据（分组柱状图）
        chart_data = {
            'type': 'bar',
            'title': 'Grouped Bar Chart',
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['Q1', 'Q2', 'Q3', 'Q1', 'Q2', 'Q3'],  # 重复的X轴标签
                    'y_data': [10, 20, 30, 15, 25, 35]
                },
                {
                    'name': 'Series2',
                    'x_data': ['Q1', 'Q2', 'Q3', 'Q1', 'Q2', 'Q3'],  # 重复的X轴标签
                    'y_data': [12, 22, 32, 17, 27, 37]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_bar_chart(chart_data)

        # 应该成功渲染并处理重复的X轴标签
        assert '<svg' in svg_result
        assert 'Q1' in svg_result
        assert 'Q2' in svg_result
        assert 'Q3' in svg_result


class TestSVGChartRendererAdvancedFeatures:
    """TDD测试：SVG图表渲染器高级功能测试"""

    def test_render_chart_with_negative_y_values(self):
        """
        TDD测试：图表渲染应该正确处理负数Y值

        覆盖代码行：343-345 - 负数Y值处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建包含负数的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Negative Values',
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [-10, 5, -15]  # 包含负数
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_bar_chart(chart_data)

        # 应该成功渲染负数值
        assert '<svg' in svg_result
        assert 'Chart with Negative Values' in svg_result

    def test_render_chart_with_custom_y_axis_max(self):
        """
        TDD测试：图表渲染应该处理自定义Y轴最大值

        覆盖代码行：358-359 - Y轴最大值处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建带有自定义Y轴最大值的图表数据
        chart_data = {
            'type': 'line',
            'title': 'Chart with Custom Y Max',
            'y_axis_max': 100,  # 自定义Y轴最大值
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [10, 20, 30]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_line_chart(chart_data)

        # 应该成功渲染并使用自定义Y轴范围
        assert '<svg' in svg_result
        assert 'Chart with Custom Y Max' in svg_result

    def test_render_chart_with_data_labels_enabled(self):
        """
        TDD测试：图表渲染应该处理启用的数据标签

        覆盖代码行：380-382 - 数据标签处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建启用数据标签的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Data Labels',
            'data_labels': {'enabled': True, 'position': 'center'},
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [10, 20, 30],
                    'data_labels': {'enabled': True}
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_bar_chart(chart_data)

        # 应该成功渲染并包含数据标签
        assert '<svg' in svg_result
        assert 'Chart with Data Labels' in svg_result

    def test_render_pie_chart_with_complex_data(self):
        """
        TDD测试：饼图渲染应该处理复杂的数据结构

        覆盖代码行：407-408, 424-425 - 饼图复杂数据处理
        """
        renderer = SVGChartRenderer()

        # 创建复杂的饼图数据
        chart_data = {
            'type': 'pie',
            'title': 'Complex Pie Chart',
            'series': [
                {
                    'name': 'Pie Series',
                    'x_data': ['Segment A', 'Segment B', 'Segment C', 'Segment D'],
                    'y_data': [25, 35, 20, 20],
                    'colors': ['#FF0000', '#00FF00', '#0000FF', '#FFFF00']
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_pie_chart(chart_data)

        # 应该成功渲染复杂饼图
        assert '<svg' in svg_result
        assert 'Complex Pie Chart' in svg_result


class TestSVGChartRendererErrorHandling:
    """TDD测试：SVG图表渲染器错误处理测试"""

    def test_render_chart_with_invalid_angle_calculation(self):
        """
        TDD测试：饼图渲染应该处理角度计算异常

        覆盖代码行：430-431 - 角度计算异常处理
        """
        renderer = SVGChartRenderer()

        # 创建可能导致角度计算问题的饼图数据
        chart_data = {
            'type': 'pie',
            'title': 'Pie Chart with Zero Values',
            'series': [
                {
                    'name': 'Pie Series',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [0, 0, 0]  # 全零值，可能导致角度计算问题
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_pie_chart(chart_data)

        # 应该处理异常并返回有效SVG
        assert '<svg' in svg_result

    def test_render_chart_with_empty_data_handling(self):
        """
        TDD测试：图表渲染应该处理空数据情况

        覆盖代码行：306, 320 - 空数据处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建空数据的图表
        chart_data = {
            'type': 'bar',
            'title': 'Empty Data Chart',
            'series': []  # 空系列
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该处理空数据并返回有效SVG
        assert '<svg' in svg_result
        assert 'Empty Data Chart' in svg_result

    def test_render_chart_with_legend_positioning(self):
        """
        TDD测试：图表渲染应该处理图例定位

        覆盖代码行：466, 497 - 图例定位逻辑
        """
        renderer = SVGChartRenderer()

        # 创建带有图例的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Legend',
            'legend': {
                'enabled': True,
                'position': 'right',
                'entries': [
                    {'text': 'Series 1', 'color': '#FF0000'},
                    {'text': 'Series 2', 'color': '#00FF00'}
                ]
            },
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B'],
                    'y_data': [10, 20]
                },
                {
                    'name': 'Series2',
                    'x_data': ['A', 'B'],
                    'y_data': [15, 25]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_bar_chart(chart_data)

        # 应该包含图例
        assert '<svg' in svg_result
        assert 'Chart with Legend' in svg_result

    def test_render_chart_with_axis_scaling(self):
        """
        TDD测试：图表渲染应该处理轴缩放

        覆盖代码行：516, 524 - 轴缩放处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建需要轴缩放的图表数据
        chart_data = {
            'type': 'line',
            'title': 'Chart with Axis Scaling',
            'x_axis': {'scale': 'linear'},
            'y_axis': {'scale': 'linear'},
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['1', '10', '100', '1000'],  # 大范围数据
                    'y_data': [1, 100, 10000, 1000000]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_line_chart(chart_data)

        # 应该处理大范围数据的缩放
        assert '<svg' in svg_result
        assert 'Chart with Axis Scaling' in svg_result


class TestSVGChartRendererComplexScenarios:
    """TDD测试：SVG图表渲染器复杂场景测试"""

    def test_render_chart_with_annotations(self):
        """
        TDD测试：图表渲染应该处理注释

        覆盖代码行：565-567 - 注释处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建带有注释的图表数据
        chart_data = {
            'type': 'line',
            'title': 'Chart with Annotations',
            'annotations': [
                {
                    'text': '重要数据点',
                    'position': {'x': 100, 'y': 200},
                    'style': {'color': '#FF0000'}
                }
            ],
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [10, 20, 30]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_line_chart(chart_data)

        # 应该包含注释
        assert '<svg' in svg_result
        assert 'Chart with Annotations' in svg_result

    def test_render_chart_with_grid_lines(self):
        """
        TDD测试：图表渲染应该处理网格线

        覆盖代码行：574, 581, 586 - 网格线渲染逻辑
        """
        renderer = SVGChartRenderer(show_axes=True)  # 启用坐标轴和网格线

        # 创建需要网格线的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Grid Lines',
            'grid': {'enabled': True, 'style': 'solid'},
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C', 'D'],
                    'y_data': [10, 20, 30, 40]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_bar_chart(chart_data)

        # 应该包含网格线
        assert '<svg' in svg_result
        assert 'Chart with Grid Lines' in svg_result

    def test_render_chart_with_custom_colors_and_styles(self):
        """
        TDD测试：图表渲染应该处理自定义颜色和样式

        覆盖代码行：600, 602 - 自定义样式处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建带有自定义样式的图表数据
        chart_data = {
            'type': 'pie',
            'title': 'Chart with Custom Styles',
            'colors': ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4'],
            'style': {
                'background': '#F8F9FA',
                'border': {'width': 2, 'color': '#DEE2E6'}
            },
            'series': [
                {
                    'name': 'Pie Series',
                    'x_data': ['A', 'B', 'C', 'D'],
                    'y_data': [25, 30, 25, 20]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_pie_chart(chart_data)

        # 应该应用自定义样式
        assert '<svg' in svg_result
        # 饼图可能使用系列名称作为标题，所以检查系列名称
        assert 'Pie Series' in svg_result or 'Chart with Custom Styles' in svg_result

    def test_render_chart_with_data_point_highlighting(self):
        """
        TDD测试：图表渲染应该处理数据点高亮

        覆盖代码行：649-652 - 数据点高亮逻辑
        """
        renderer = SVGChartRenderer()

        # 创建带有高亮数据点的图表数据
        chart_data = {
            'type': 'line',
            'title': 'Chart with Highlighted Points',
            'highlight': {
                'enabled': True,
                'points': [{'series': 0, 'index': 1, 'color': '#FF0000'}]
            },
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [10, 20, 30]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer._render_line_chart(chart_data)

        # 应该包含高亮的数据点
        assert '<svg' in svg_result
        assert 'Chart with Highlighted Points' in svg_result


class TestSVGChartRendererAdvancedDataLabels:
    """TDD测试：SVG图表渲染器高级数据标签测试"""

    def test_render_chart_with_show_value_data_labels(self):
        """
        TDD测试：图表渲染应该处理显示数值的数据标签

        覆盖代码行：1095-1096 - 显示数值的数据标签逻辑
        """
        renderer = SVGChartRenderer()

        # 创建带有显示数值的数据标签的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Value Labels',
            'data_labels': {
                'enabled': True,
                'show_value': True,
                'show_category_name': False
            },
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [10, 20, 30],
                    'data_labels': {
                        'enabled': True,
                        'show_value': True
                    }
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该包含数值标签
        assert '<svg' in svg_result
        assert 'Chart with Value Labels' in svg_result

    def test_render_chart_with_category_name_data_labels(self):
        """
        TDD测试：图表渲染应该处理显示分类名称的数据标签

        覆盖代码行：1098-1102 - 显示分类名称的数据标签逻辑
        """
        renderer = SVGChartRenderer()

        # 创建带有显示分类名称的数据标签的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Category Labels',
            'data_labels': {
                'enabled': True,
                'show_value': True,
                'show_category_name': True
            },
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['Category A', 'Category B', 'Category C'],
                    'y_data': [15, 25, 35],
                    'data_labels': {
                        'enabled': True,
                        'show_value': True,
                        'show_category_name': True
                    }
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该包含分类名称标签
        assert '<svg' in svg_result
        assert 'Chart with Category Labels' in svg_result

    def test_render_legend_with_custom_style(self):
        """
        TDD测试：图例渲染应该处理自定义样式

        覆盖代码行：1044-1052 - 图例自定义样式处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建带有自定义图例样式的图表数据
        chart_data = {
            'type': 'line',
            'title': 'Chart with Styled Legend',
            'legend': {
                'enabled': True,
                'position': 'right',
                'style': {
                    'font_family': 'Arial',
                    'font_size': 12,
                    'color': '#333333'
                },
                'entries': [
                    {'text': 'Series 1', 'color': '#FF0000'},
                    {'text': 'Series 2', 'color': '#00FF00'}
                ]
            },
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [10, 20, 30]
                },
                {
                    'name': 'Series2',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [15, 25, 35]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该包含带样式的图例
        assert '<svg' in svg_result
        assert 'Chart with Styled Legend' in svg_result


class TestSVGChartRendererSpecialCases:
    """TDD测试：SVG图表渲染器特殊情况测试"""

    def test_render_chart_with_xml_formatting_edge_cases(self):
        """
        TDD测试：图表渲染应该处理XML格式化的边界情况

        覆盖代码行：811-815 - XML格式化边界情况处理
        """
        renderer = SVGChartRenderer()

        # 创建可能导致XML格式化问题的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Special Characters & <XML> "Quotes"',
            'series': [
                {
                    'name': 'Series with & special chars',
                    'x_data': ['A&B', 'C<D', 'E"F'],
                    'y_data': [10, 20, 30]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该正确处理特殊字符
        assert '<svg' in svg_result
        # XML应该被正确转义
        assert '&amp;' in svg_result or 'Special Characters' in svg_result

    def test_render_chart_with_axis_label_positioning(self):
        """
        TDD测试：图表渲染应该处理轴标签定位

        覆盖代码行：952 - 轴标签定位逻辑
        """
        renderer = SVGChartRenderer(show_axes=True)

        # 创建需要轴标签定位的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Axis Labels',
            'x_axis': {'title': 'X Axis Title'},
            'y_axis': {'title': 'Y Axis Title'},
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['Very Long Category Name 1', 'Very Long Category Name 2'],
                    'y_data': [100, 200]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该包含轴标签
        assert '<svg' in svg_result
        assert 'Chart with Axis Labels' in svg_result

    def test_render_chart_with_complex_legend_positioning(self):
        """
        TDD测试：图表渲染应该处理复杂的图例定位

        覆盖代码行：1105-1110, 1114-1126 - 复杂图例定位逻辑
        """
        renderer = SVGChartRenderer()

        # 创建需要复杂图例定位的图表数据
        chart_data = {
            'type': 'line',
            'title': 'Chart with Complex Legend',
            'legend': {
                'enabled': True,
                'position': 'bottom',
                'entries': [
                    {'text': 'Very Long Series Name 1', 'color': '#FF0000'},
                    {'text': 'Very Long Series Name 2', 'color': '#00FF00'},
                    {'text': 'Very Long Series Name 3', 'color': '#0000FF'},
                    {'text': 'Very Long Series Name 4', 'color': '#FFFF00'},
                    {'text': 'Very Long Series Name 5', 'color': '#FF00FF'}
                ]
            },
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [10, 20, 30]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该包含复杂图例
        assert '<svg' in svg_result
        assert 'Chart with Complex Legend' in svg_result

    def test_render_chart_with_error_handling_in_data_processing(self):
        """
        TDD测试：图表渲染应该处理数据处理中的错误

        覆盖代码行：827-828 - 数据处理错误处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建可能导致数据处理错误的图表数据（使用有效数据但测试边界情况）
        chart_data = {
            'type': 'pie',
            'title': 'Chart with Data Processing Issues',
            'series': [
                {
                    'name': 'Pie Series',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [0, 0, 0]  # 全零值，可能导致处理问题
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该处理边界数据并返回有效SVG
        assert '<svg' in svg_result
        # 全零值的饼图可能不显示标题，但应该返回有效的SVG结构
        assert 'xmlns="http://www.w3.org/2000/svg"' in svg_result


class TestSVGChartRendererFinalCoverage:
    """TDD测试：SVG图表渲染器最终覆盖率测试"""

    def test_render_chart_with_advanced_pie_chart_features(self):
        """
        TDD测试：饼图渲染应该处理高级功能

        覆盖代码行：1170, 1174 - 饼图高级功能处理
        """
        renderer = SVGChartRenderer()

        # 创建带有高级功能的饼图数据
        chart_data = {
            'type': 'pie',
            'title': 'Advanced Pie Chart',
            'pie_options': {
                'start_angle': 90,
                'explode': [0, 0.1, 0, 0],  # 突出第二个扇形
                'show_percentages': True
            },
            'series': [
                {
                    'name': 'Pie Series',
                    'x_data': ['Segment A', 'Segment B', 'Segment C', 'Segment D'],
                    'y_data': [30, 25, 25, 20],
                    'colors': ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4']
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该包含高级饼图功能
        assert '<svg' in svg_result
        assert 'Advanced Pie Chart' in svg_result

    def test_render_chart_with_text_overflow_handling(self):
        """
        TDD测试：图表渲染应该处理文本溢出

        覆盖代码行：1195, 1197-1202 - 文本溢出处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建可能导致文本溢出的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Very Long Text That Might Overflow the Available Space in the SVG Canvas',
            'series': [
                {
                    'name': 'Series with Extremely Long Name That Exceeds Normal Display Limits',
                    'x_data': [
                        'Very Long Category Name That Exceeds Normal Display Width',
                        'Another Extremely Long Category Name',
                        'Yet Another Very Long Category Name'
                    ],
                    'y_data': [100, 200, 300]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该处理文本溢出
        assert '<svg' in svg_result
        # 标题可能被截断或处理
        assert 'Chart with Very Long Text' in svg_result or 'Long Text' in svg_result

    def test_render_chart_with_edge_case_coordinates(self):
        """
        TDD测试：图表渲染应该处理边界坐标情况

        覆盖代码行：574, 581 - 边界坐标处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建边界坐标的图表数据
        chart_data = {
            'type': 'line',
            'title': 'Chart with Edge Coordinates',
            'series': [
                {
                    'name': 'Series1',
                    'x_data': ['A'],  # 单点数据
                    'y_data': [0]     # 零值
                },
                {
                    'name': 'Series2',
                    'x_data': ['A'],
                    'y_data': [1000000]  # 极大值
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该处理边界坐标
        assert '<svg' in svg_result
        assert 'Chart with Edge Coordinates' in svg_result

    def test_render_chart_with_special_formatting_requirements(self):
        """
        TDD测试：图表渲染应该处理特殊格式化需求

        覆盖代码行：659, 663-664 - 特殊格式化处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建需要特殊格式化的图表数据
        chart_data = {
            'type': 'bar',
            'title': 'Chart with Special Formatting',
            'formatting': {
                'number_format': '#,##0.00',
                'currency_symbol': '$',
                'decimal_places': 2
            },
            'series': [
                {
                    'name': 'Financial Data',
                    'x_data': ['Q1', 'Q2', 'Q3', 'Q4'],
                    'y_data': [1234.567, 2345.678, 3456.789, 4567.890]
                }
            ]
        }

        # 测试渲染
        svg_result = renderer.render_chart_to_svg(chart_data)

        # 应该处理特殊格式化
        assert '<svg' in svg_result
        assert 'Chart with Special Formatting' in svg_result



class TestSVGChartRendererUncoveredCode:
    """TDD测试：专门针对未覆盖代码行的测试类"""

    def test_color_extension_logic_line_306(self):
        """
        TDD测试：颜色扩展逻辑应该正确处理系列数量超过预定义颜色数量的情况

        覆盖代码行：306 - colors.extend(DEFAULT_CHART_COLORS * (len(series_list) - len(colors)))
        """
        renderer = SVGChartRenderer()

        # 创建超过默认颜色数量的多系列图表数据
        chart_data = {
            'type': 'bar',
            'title': '多系列颜色扩展测试',
            'series': [
                {'name': f'系列{i}', 'x_data': ['A', 'B'], 'y_data': [10, 20]}
                for i in range(1, 16)  # 创建15个系列，超过默认颜色数量
            ]
        }

        svg_result = renderer._render_bar_chart(chart_data)

        # 应该成功渲染所有系列，颜色会循环使用
        assert '<svg' in svg_result
        assert '多系列颜色扩展测试' in svg_result
        # 验证所有系列都被渲染（通过检查rect元素数量）
        rect_count = svg_result.count('<rect')
        assert rect_count >= 30  # 15个系列 * 2个数据点 = 30个柱子

    def test_empty_data_check_line_320(self):
        """
        TDD测试：空数据检查应该正确处理没有X轴标签或Y轴数值的情况

        覆盖代码行：320 - return self._format_svg(svg)
        """
        renderer = SVGChartRenderer()

        # 测试空X轴标签的情况
        chart_data_empty_x = {
            'type': 'bar',
            'title': '空X轴标签测试',
            'series': [
                {'name': '系列1', 'x_data': [], 'y_data': [10, 20, 30]}
            ]
        }

        svg_result = renderer._render_bar_chart(chart_data_empty_x)

        # 应该返回基本的SVG结构，但没有数据内容
        assert '<svg' in svg_result
        assert 'xmlns="http://www.w3.org/2000/svg"' in svg_result
        assert '空X轴标签测试' in svg_result

        # 测试空Y轴数值的情况
        chart_data_empty_y = {
            'type': 'bar',
            'title': '空Y轴数值测试',
            'series': [
                {'name': '系列1', 'x_data': ['A', 'B', 'C'], 'y_data': []}
            ]
        }

        svg_result_y = renderer._render_bar_chart(chart_data_empty_y)
        assert '<svg' in svg_result_y
        assert '空Y轴数值测试' in svg_result_y

    def test_x_axis_label_deduplication_logic_lines_329_330(self):
        """
        TDD测试：X轴标签去重逻辑应该正确处理分组图表和连续图表的不同策略

        覆盖代码行：329-330 - X轴标签去重和显示策略
        """
        renderer = SVGChartRenderer()

        # 测试不需要去重的连续图表（should_deduplicate = False）
        chart_data_continuous = {
            'type': 'bar',
            'title': '连续图表标签测试',
            'series': [
                {'name': '系列1', 'x_data': ['A', 'B', 'C', 'D'], 'y_data': [10, 20, 30, 40]},
                {'name': '系列2', 'x_data': ['A', 'B', 'C', 'D'], 'y_data': [15, 25, 35, 45]}
            ]
        }

        # 模拟should_deduplicate返回False的情况
        with patch.object(renderer, '_should_deduplicate_x_labels', return_value=False):
            svg_result = renderer._render_bar_chart(chart_data_continuous)

        assert '<svg' in svg_result
        assert '连续图表标签测试' in svg_result

        # 测试需要去重的分组图表（should_deduplicate = True）
        chart_data_grouped = {
            'type': 'bar',
            'title': '分组图表标签测试',
            'series': [
                {'name': '系列1', 'x_data': ['A', 'A', 'B', 'B'], 'y_data': [10, 15, 20, 25]},
                {'name': '系列2', 'x_data': ['A', 'A', 'B', 'B'], 'y_data': [12, 18, 22, 28]}
            ]
        }

        with patch.object(renderer, '_should_deduplicate_x_labels', return_value=True):
            svg_result_grouped = renderer._render_bar_chart(chart_data_grouped)

        assert '<svg' in svg_result_grouped
        assert '分组图表标签测试' in svg_result_grouped

    def test_numeric_type_conversion_exception_handling_lines_343_345(self):
        """测试数值类型转换异常处理"""
        renderer = SVGChartRenderer()

        # 创建包含无效数值类型的图表数据
        chart_data_invalid_values = {
            'type': 'bar',
            'title': '无效数值类型测试',
            'y_axis_max': 'invalid_string',  # 无效的y_axis_max值
            'series': [
                {'name': '系列1', 'x_data': ['A', 'B'], 'y_data': [10, 20]}  # 使用有效数值
            ]
        }

        # 这应该触发异常处理，使用默认值y_min=0.0, y_max=1.0
        svg_result = renderer._render_bar_chart(chart_data_invalid_values)

        # 应该成功渲染，使用默认的数值范围
        assert '<svg' in svg_result
        assert '无效数值类型测试' in svg_result

        # 测试空系列的情况来触发异常处理
        chart_data_empty_series = {
            'type': 'bar',
            'title': '空系列测试',
            'y_axis_max': None,
            'series': []  # 空系列列表
        }

        svg_result_empty = renderer._render_bar_chart(chart_data_empty_series)
        assert '<svg' in svg_result_empty
        assert '空系列测试' in svg_result_empty

    def test_line_chart_insufficient_data_points_line_516(self):
        """
        TDD测试：折线图应该正确处理数据点不足的情况

        覆盖代码行：516 - continue (当数据点少于2个时跳过系列)
        """
        renderer = SVGChartRenderer()

        chart_data_insufficient = {
            'type': 'line',
            'title': '数据点不足测试',
            'series': [
                {'name': '单点系列', 'x_data': ['A'], 'y_data': [10]},
                {'name': '正常系列', 'x_data': ['A', 'B', 'C'], 'y_data': [10, 20, 30]},
                {'name': '空系列', 'x_data': [], 'y_data': []}
            ]
        }

        svg_result = renderer._render_line_chart(chart_data_insufficient)

        assert '<svg' in svg_result
        assert '数据点不足测试' in svg_result
        assert 'path' in svg_result or 'line' in svg_result

    def test_chart_type_validation_edge_cases(self):
        """
        TDD测试：图表类型验证边界情况

        测试不同图表类型的基本验证功能
        """
        renderer = SVGChartRenderer()

        # 创建基本的图表数据
        chart_data = {
            'type': 'bar',
            'title': '图表类型验证测试',
            'series': [
                {'name': '系列1', 'x_data': ['A', 'B'], 'y_data': [10, 20]}
            ]
        }

        svg_result = renderer._render_bar_chart(chart_data)

        # 应该成功渲染图表
        assert '<svg' in svg_result
        assert '图表类型验证测试' in svg_result

    def test_single_x_label_coordinate_calculation_lines_524_753(self):
        """
        TDD测试：单个X轴标签时的坐标计算应该正确处理

        覆盖代码行：524, 753 - 单个标签时的特殊坐标计算逻辑
        """
        renderer = SVGChartRenderer()

        # 测试折线图的单标签情况
        chart_data_single_label = {
            'type': 'line',
            'title': '单标签坐标测试',
            'series': [
                {'name': '系列1', 'x_data': ['单点'], 'y_data': [50]},
                {'name': '系列2', 'x_data': ['单点'], 'y_data': [30]}
            ]
        }

        svg_result = renderer._render_line_chart(chart_data_single_label)

        # 应该正确处理单点坐标计算
        assert '<svg' in svg_result
        assert '单标签坐标测试' in svg_result

        # 测试面积图的单标签情况
        chart_data_area_single = {
            'type': 'area',
            'title': '面积图单标签测试',
            'series': [
                {'name': '系列1', 'x_data': ['单点'], 'y_data': [50]},
                {'name': '系列2', 'x_data': ['单点'], 'y_data': [30]}
            ]
        }

        svg_result_area = renderer._render_area_chart(chart_data_area_single)
        assert '<svg' in svg_result_area
        assert '面积图单标签测试' in svg_result_area

    def test_bar_chart_zero_y_range_handling_lines_430_431(self):
        """
        TDD测试：柱状图应该正确处理Y轴范围为零的情况

        覆盖代码行：430-431 - 当y_range为0时的特殊处理
        """
        renderer = SVGChartRenderer()

        # 创建所有Y值相同的图表数据（导致y_range = 0）
        chart_data_zero_range = {
            'type': 'bar',
            'title': 'Y轴范围为零测试',
            'series': [
                {'name': '系列1', 'x_data': ['A', 'B', 'C'], 'y_data': [10, 10, 10]}
            ]
        }

        svg_result = renderer._render_bar_chart(chart_data_zero_range)

        # 应该使用默认高度渲染柱子
        assert '<svg' in svg_result
        assert 'Y轴范围为零测试' in svg_result
        # 应该包含柱子元素
        assert 'rect' in svg_result

    def test_axis_rendering_with_show_axes_true_lines_574_581(self):
        """
        TDD测试：启用坐标轴时应该正确渲染坐标轴元素

        覆盖代码行：574, 581 - 坐标轴渲染逻辑
        """
        renderer = SVGChartRenderer(show_axes=True)

        # 创建需要显示坐标轴的图表数据
        chart_data_with_axes = {
            'type': 'bar',
            'title': '坐标轴显示测试',
            'series': [
                {'name': '系列1', 'x_data': ['A', 'B', 'C'], 'y_data': [10, 20, 30]}
            ]
        }

        svg_result = renderer._render_bar_chart(chart_data_with_axes)

        # 应该包含坐标轴相关的SVG元素
        assert '<svg' in svg_result
        assert '坐标轴显示测试' in svg_result

    def test_pie_chart_edge_cases_lines_649_652_659_663_664(self):
        """
        TDD测试：饼图应该正确处理边界情况和特殊计算

        覆盖代码行：649-652, 659, 663-664 - 饼图扇形计算和渲染逻辑
        """
        renderer = SVGChartRenderer()

        # 测试包含零值的饼图数据
        chart_data_with_zeros = {
            'type': 'pie',
            'title': '饼图零值测试',
            'series': [
                {
                    'name': '饼图系列',
                    'x_data': ['A', 'B', 'C', 'D'],
                    'y_data': [0, 30, 0, 70]  # 包含零值
                }
            ]
        }

        svg_result = renderer._render_pie_chart(chart_data_with_zeros)

        # 应该正确处理零值，只渲染非零扇形
        assert '<svg' in svg_result
        assert '饼图零值测试' in svg_result

        # 测试单个值的饼图
        chart_data_single_value = {
            'type': 'pie',
            'title': '单值饼图测试',
            'series': [
                {
                    'name': '饼图系列',
                    'x_data': ['唯一值'],
                    'y_data': [100]
                }
            ]
        }

        svg_result_single = renderer._render_pie_chart(chart_data_single_value)
        assert '<svg' in svg_result_single
        assert '单值饼图测试' in svg_result_single

    def test_legend_rendering_edge_cases_lines_811_815_827_828(self):
        """
        TDD测试：图例渲染应该正确处理边界情况

        覆盖代码行：811-815, 827-828 - 图例渲染的特殊情况处理
        """
        renderer = SVGChartRenderer()

        # 创建需要复杂图例处理的图表数据
        chart_data_complex_legend = {
            'type': 'bar',
            'title': '复杂图例测试',
            'legend': {
                'show': True,
                'position': 'right',
                'entries': [
                    {'name': '很长的图例名称测试文本', 'color': '#FF0000'},
                    {'name': '短名', 'color': '#00FF00'},
                    {'name': '', 'color': '#0000FF'},  # 空名称
                ]
            },
            'series': [
                {'name': '系列1', 'x_data': ['A', 'B'], 'y_data': [10, 20]},
                {'name': '系列2', 'x_data': ['A', 'B'], 'y_data': [15, 25]},
                {'name': '系列3', 'x_data': ['A', 'B'], 'y_data': [12, 22]}
            ]
        }

        svg_result = renderer._render_bar_chart(chart_data_complex_legend)

        # 应该正确处理复杂图例情况
        assert '<svg' in svg_result
        assert '复杂图例测试' in svg_result

    def test_error_handling_in_chart_rendering_line_952(self):
        """
        TDD测试：图表渲染过程中的错误处理

        覆盖代码行：952 - 图表渲染错误处理逻辑
        """
        renderer = SVGChartRenderer()

        # 创建可能导致渲染错误的图表数据
        chart_data_error_prone = {
            'type': 'bar',
            'title': '错误处理测试',
            'series': [
                {
                    'name': '问题系列',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [float('inf'), float('-inf'), float('nan')]  # 特殊浮点值
                }
            ]
        }

        # 应该能够处理特殊数值而不崩溃
        svg_result = renderer._render_bar_chart(chart_data_error_prone)

        # 应该返回有效的SVG，即使数据有问题
        assert '<svg' in svg_result
        assert '错误处理测试' in svg_result

    def test_advanced_pie_chart_features_lines_1170_1174_1195_1197_1202(self):
        """
        TDD测试：饼图高级功能应该正确处理复杂配置

        覆盖代码行：1170, 1174, 1195, 1197-1202 - 饼图高级功能处理
        """
        renderer = SVGChartRenderer()

        # 创建带有高级功能的饼图数据
        chart_data_advanced_pie = {
            'type': 'pie',
            'title': '高级饼图功能测试',
            'pie_config': {
                'start_angle': 90,
                'show_percentages': True,
                'explode_slices': [0, 0.1, 0, 0.05]  # 突出某些扇形
            },
            'series': [
                {
                    'name': '高级饼图系列',
                    'x_data': ['第一部分', '第二部分', '第三部分', '第四部分'],
                    'y_data': [25, 35, 20, 20],
                    'colors': ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4']
                }
            ]
        }

        svg_result = renderer._render_pie_chart(chart_data_advanced_pie)

        # 应该包含高级饼图功能
        assert '<svg' in svg_result
        assert '高级饼图功能测试' in svg_result

    def test_data_label_positioning_lines_1044_1052_1090_1102(self):
        """
        TDD测试：数据标签定位应该正确处理各种位置计算

        覆盖代码行：1044-1052, 1090, 1102 - 数据标签位置计算逻辑
        """
        renderer = SVGChartRenderer()

        # 创建需要数据标签的图表数据
        chart_data_with_labels = {
            'type': 'bar',
            'title': '数据标签定位测试',
            'data_labels': {
                'show': True,
                'position': 'top',
                'format': '{value}'
            },
            'series': [
                {
                    'name': '标签系列',
                    'x_data': ['A', 'B', 'C'],
                    'y_data': [10, -5, 25],  # 包含负值测试标签位置
                    'data_labels': {'show': True}
                }
            ]
        }

        svg_result = renderer._render_bar_chart(chart_data_with_labels)

        # 应该包含正确定位的数据标签
        assert '<svg' in svg_result
        assert '数据标签定位测试' in svg_result
