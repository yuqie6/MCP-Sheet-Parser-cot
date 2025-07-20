
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


@patch('src.converters.chart_converter.SVGChartRenderer')
def test_render_chart_content_with_exception(mock_renderer, chart_converter, sample_chart):
    """
    TDD测试：_render_chart_content应该处理渲染异常

    这个测试覆盖第29-31行的异常处理代码路径
    """
    mock_renderer.return_value.render_chart_to_svg.side_effect = Exception("Rendering failed")

    with patch('src.converters.chart_converter.logger') as mock_logger:
        html = chart_converter._render_chart_content(sample_chart)

        # 应该记录警告并返回错误信息
        mock_logger.warning.assert_called_once()
        assert "Chart rendering failed" in html
        assert "chart-error" in html

def test_generate_overlay_charts_html_with_no_charts():
    """
    TDD测试：generate_overlay_charts_html应该处理没有图表的工作表

    这个测试覆盖第50行的边界情况
    """
    chart_converter = ChartConverter(MagicMock())
    sheet = Sheet(name="EmptySheet", rows=[], charts=[])

    html = chart_converter.generate_overlay_charts_html(sheet)

    # 应该返回空字符串或最小的HTML结构
    assert html == "" or "<div" in html

def test_generate_overlay_charts_html_with_chart_without_position():
    """
    TDD测试：generate_overlay_charts_html应该处理没有位置信息的图表

    这个测试确保方法在图表没有位置信息时正确处理
    """
    chart_converter = ChartConverter(MagicMock())

    chart = Chart(name="NoPositionChart", type="bar", anchor="A1")
    chart.chart_data = {"type": "bar", "series": []}
    chart.position = None  # 没有位置信息

    sheet = Sheet(name="Sheet1", rows=[], charts=[chart])

    with patch('src.converters.chart_converter.create_position_calculator') as mock_pos_calc:
        mock_calculator = MagicMock()
        mock_calculator.calculate_chart_css_position.return_value = None
        mock_pos_calc.return_value = mock_calculator

        html = chart_converter.generate_overlay_charts_html(sheet)

        # 应该能够处理没有位置的情况
        assert isinstance(html, str)

def test_generate_standalone_charts_html_with_empty_list():
    """
    TDD测试：generate_standalone_charts_html应该处理空图表列表

    这个测试确保方法在没有图表时的行为
    """
    chart_converter = ChartConverter(MagicMock())

    html = chart_converter.generate_standalone_charts_html([])

    # 应该返回包含标题但没有图表内容的HTML
    assert "<h2>Charts</h2>" in html

def test_generate_standalone_charts_html_with_multiple_charts():
    """
    TDD测试：generate_standalone_charts_html应该处理多个图表

    这个测试验证方法能正确处理多个图表
    """
    chart_converter = ChartConverter(MagicMock())

    chart1 = Chart(name="Chart1", type="bar", anchor="A1")
    chart1.chart_data = {"type": "bar", "series": []}

    chart2 = Chart(name="Chart2", type="line", anchor="C1")
    chart2.chart_data = {"type": "line", "series": []}

    html = chart_converter.generate_standalone_charts_html([chart1, chart2])

    # 应该包含两个图表的标题
    assert "<h3>Chart1</h3>" in html
    assert "<h3>Chart2</h3>" in html
    assert html.count("<h3>") == 2

@patch('src.converters.chart_converter.SVGChartRenderer')
def test_render_chart_content_with_chart_data_none(mock_renderer, chart_converter):
    """
    TDD测试：_render_chart_content应该处理chart_data为None的情况

    这个测试覆盖第32-33行的代码路径
    """
    chart = Chart(name="NullDataChart", type="bar", anchor="A1")
    chart.chart_data = None

    html = chart_converter._render_chart_content(chart)

    # 应该返回占位符信息
    assert "Chart data not available" in html
    assert "chart-placeholder" in html

@patch('src.converters.chart_converter.SVGChartRenderer')
def test_render_chart_content_with_empty_chart_data(mock_renderer, chart_converter):
    """
    TDD测试：_render_chart_content应该处理空的chart_data

    这个测试确保方法在chart_data为空字典时的行为
    """
    chart = Chart(name="EmptyDataChart", type="bar", anchor="A1")
    chart.chart_data = {}  # 空字典

    mock_renderer.return_value.render_chart_to_svg.return_value = "<svg>empty chart</svg>"

    html = chart_converter._render_chart_content(chart)

    # 应该尝试渲染空数据
    assert "<svg>empty chart</svg>" in html

def test_chart_converter_initialization():
    """
    TDD测试：ChartConverter应该正确初始化

    这个测试验证构造函数的正确性
    """
    mock_cell_converter = MagicMock()
    converter = ChartConverter(mock_cell_converter)

    # 验证cell_converter被正确设置
    assert converter.cell_converter is mock_cell_converter


@patch('src.converters.chart_converter.create_position_calculator')
def test_generate_overlay_charts_html_with_image_chart(mock_pos_calc, chart_converter):
    """
    TDD测试：generate_overlay_charts_html应该正确处理图片类型的图表

    这个测试覆盖第50行的图片类型高度计算代码
    """
    # 创建一个图片类型的图表
    image_chart = Chart(name="ImageChart", type="image", anchor="A1")
    image_chart.chart_data = {"type": "image", "data": "image_data"}
    image_chart.position = ChartPosition(
        from_col=0, from_row=0, to_col=5, to_row=10,
        from_col_offset=0, from_row_offset=0, to_col_offset=0, to_row_offset=0
    )

    sheet = Sheet(name="Sheet1", rows=[], charts=[image_chart])

    # 模拟位置计算器
    mock_calculator = MagicMock()
    css_pos = MagicMock()
    css_pos.width = 300
    css_pos.height = 200  # 这将触发图片高度计算：max(150, int(200 * 1.333)) = 266
    mock_calculator.calculate_chart_css_position.return_value = css_pos
    mock_calculator.generate_chart_html_with_positioning.return_value = "<div>positioned image chart</div>"
    mock_pos_calc.return_value = mock_calculator

    # 模拟_render_chart_content方法
    with patch.object(chart_converter, '_render_chart_content', return_value="<img>chart content</img>") as mock_render:
        html = chart_converter.generate_overlay_charts_html(sheet)

        # 验证返回了定位的HTML
        assert "<div>positioned image chart</div>" in html

        # 验证_render_chart_content被调用时使用了正确的参数
        mock_render.assert_called_once()
        call_args = mock_render.call_args[0]
        assert call_args[0] == image_chart  # 图表对象
        assert call_args[1] == 300  # 宽度
        assert call_args[2] == 266  # 高度（200 * 1.333 = 266.6，取整为266）
