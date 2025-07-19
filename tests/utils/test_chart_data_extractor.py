import pytest
from unittest.mock import MagicMock, PropertyMock, patch
from src.utils.chart_data_extractor import ChartDataExtractor

@pytest.fixture
def extractor():
    """提供一个 ChartDataExtractor 实例。"""
    return ChartDataExtractor()

def test_extract_series_y_data(extractor):
    """测试 extract_series_y_data 方法。"""
    series = MagicMock()
    series.val.numRef.numCache.pt = [MagicMock(v=10.5), MagicMock(v=20.0)]
    assert extractor.extract_series_y_data(series) == [10.5, 20.0]

def test_extract_series_x_data(extractor):
    """测试 extract_series_x_data 方法。"""
    series = MagicMock()
    series.cat.strRef.strCache.pt = [MagicMock(v='A'), MagicMock(v='B')]
    assert extractor.extract_series_x_data(series) == ['A', 'B']

def test_extract_axis_title_from_rich_text(extractor):
    """测试从富文本中提取轴标题。"""
    title_obj = MagicMock()
    run = MagicMock()
    type(run).t = PropertyMock(return_value='Axis Title')
    p = MagicMock()
    p.r = [run]
    type(title_obj.tx.rich).p = PropertyMock(return_value=[p])
    title_obj.tx.strRef = None 
    assert extractor.extract_axis_title(title_obj) == 'Axis Title'

def test_extract_axis_title_from_str_ref(extractor):
    """测试从字符串引用中提取轴标题。"""
    title_obj = MagicMock()
    pt = MagicMock()
    type(pt).v = PropertyMock(return_value='Cached Title')
    type(title_obj.tx.strRef.strCache).pt = PropertyMock(return_value=[pt])
    type(title_obj.tx).rich = None
    assert extractor.extract_axis_title(title_obj) == 'Cached Title'

def test_extract_axis_title_none(extractor):
    """测试当标题对象为None时的情况。"""
    assert extractor.extract_axis_title(None) is None

def test_extract_color_from_srgb(extractor):
    """测试从 srgbClr 提取颜色。"""
    solid_fill = MagicMock()
    solid_fill.srgbClr.val = 'FF0000'
    solid_fill.schemeClr = None
    assert extractor.extract_color(solid_fill) == '#FF0000'

def test_extract_color_from_scheme(extractor):
    """测试从 schemeClr 提取颜色。"""
    solid_fill = MagicMock()
    solid_fill.srgbClr = None
    solid_fill.schemeClr.val = 'accent1'
    with patch('src.utils.chart_data_extractor.convert_scheme_color_to_hex', return_value='#4472C4') as mock_convert:
        color = extractor.extract_color(solid_fill)
        mock_convert.assert_called_once_with('accent1')
        assert color == '#4472C4'

def test_extract_data_labels(extractor):
    """测试 extract_data_labels 方法。"""
    series = MagicMock()
    series.dLbls.showVal = True
    series.dLbls.showCatName = False
    labels_info = extractor.extract_data_labels(series)
    assert labels_info['enabled'] is True
    assert labels_info['show_value'] is True
    assert labels_info['show_category_name'] is False

def test_extract_legend_info(extractor):
    """测试 extract_legend_info 方法。"""
    chart = MagicMock()
    chart.legend.position = 'r'
    chart.legend.overlay = True
    legend_info = extractor.extract_legend_info(chart)
    assert legend_info['enabled'] is True
    assert legend_info['position'] == 'r'
    assert legend_info['overlay'] is True

def test_extract_legend_info_disabled(extractor):
    """测试当图例被禁用时的场景。"""
    chart = MagicMock()
    chart.legend = None
    legend_info = extractor.extract_legend_info(chart)
    assert legend_info['enabled'] is False

def test_extract_chart_title(extractor):
    """测试 extract_chart_title 方法。"""
    chart = MagicMock()
    run = MagicMock()
    type(run).t = PropertyMock(return_value='Main Chart Title')
    p = MagicMock()
    p.r = [run]
    type(chart.title.tx.rich).p = PropertyMock(return_value=[p])
    chart.title.tx.strRef = None
    assert extractor.extract_chart_title(chart) == 'Main Chart Title'

def test_extract_plot_area(extractor):
    """测试 extract_plot_area 方法。"""
    chart = MagicMock()
    spPr_mock = MagicMock()
    spPr_mock.solidFill.srgbClr.val = 'F0F0F0'
    spPr_mock.solidFill.schemeClr = None
    chart.plotArea.spPr = spPr_mock
    layout_mock = MagicMock()
    layout_mock.manualLayout.x = 0.1
    layout_mock.manualLayout.y = 0.2
    layout_mock.manualLayout.w = 0.8
    layout_mock.manualLayout.h = 0.7
    chart.plotArea.layout = layout_mock
    plot_area_info = extractor.extract_plot_area(chart)
    assert plot_area_info['background_color'] == '#F0F0F0'
    assert plot_area_info['layout']['x'] == 0.1
    assert plot_area_info['layout']['y'] == 0.2
    assert plot_area_info['layout']['width'] == 0.8
    assert plot_area_info['layout']['height'] == 0.7

def test_extract_pie_chart_colors_with_dpt(extractor):
    """测试饼图颜色提取，当数据点(dPt)有个别颜色设置时。"""
    series = MagicMock()
    dp1 = MagicMock()
    dp1.spPr.solidFill.srgbClr.val = 'FF0000'
    dp1.spPr.solidFill.schemeClr = None
    dp2 = MagicMock()
    dp2.spPr.solidFill.srgbClr = None
    dp2.spPr.solidFill.schemeClr.val = 'accent2'
    series.dPt = [dp1, dp2]
    with patch('src.utils.chart_data_extractor.convert_scheme_color_to_hex', return_value='#ED7D31') as mock_convert:
        colors = extractor.extract_pie_chart_colors(series)
        mock_convert.assert_called_once_with('accent2')
        assert colors == ['#FF0000', '#ED7D31']

def test_extract_chart_annotations_with_axes(extractor):
    """测试图表注释提取，包含坐标轴标题。"""
    chart = MagicMock()
    x_title_run = MagicMock()
    type(x_title_run).t = PropertyMock(return_value='X-Axis Label')
    x_p = MagicMock()
    x_p.r = [x_title_run]
    type(chart.x_axis.title.tx.rich).p = PropertyMock(return_value=[x_p])
    chart.x_axis.title.tx.strRef = None
    y_title_run = MagicMock()
    type(y_title_run).t = PropertyMock(return_value='Y-Axis Label')
    y_p = MagicMock()
    y_p.r = [y_title_run]
    type(chart.y_axis.title.tx.rich).p = PropertyMock(return_value=[y_p])
    chart.y_axis.title.tx.strRef = None
    chart.title = None
    with patch.object(extractor, '_extract_textbox_annotations', return_value=[]) as mock_textbox, \
         patch.object(extractor, '_extract_shape_annotations', return_value=[]) as mock_shape, \
         patch.object(extractor, '_extract_plotarea_annotations', return_value=[]) as mock_plotarea:
        annotations = extractor.extract_chart_annotations(chart)
    assert len(annotations) == 2
    assert any(a['text'] == 'X-Axis Label' and a['axis'] == 'x' for a in annotations)
    assert any(a['text'] == 'Y-Axis Label' and a['axis'] == 'y' for a in annotations)

def test_extract_from_val_get_rows_fallback(extractor):
    """测试 _extract_from_val 在 numCache 不存在时回退到 get_rows。"""
    val_obj = MagicMock()
    type(val_obj.numRef).numCache = None
    val_obj.numRef.get_rows.return_value = [MagicMock(v=30.1), MagicMock(v=40.2)]
    data = extractor._extract_from_val(val_obj)
    assert data == [30.1, 40.2]

def test_extract_from_cat_get_rows_fallback(extractor):
    """测试 _extract_from_cat 在 strCache 不存在时回退到 get_rows。"""
    cat_obj = MagicMock()
    type(cat_obj.strRef).strCache = None
    cat_obj.strRef.get_rows.return_value = [MagicMock(v='C'), MagicMock(v='D')]
    data = extractor._extract_from_cat(cat_obj)
    assert data == ['C', 'D']

def test_extract_data_labels_detailed(extractor):
    """测试 extract_data_labels 的更详细属性。"""
    series = MagicMock()
    dLbls = series.dLbls
    dLbls.showVal = True
    dLbls.showCatName = False
    dLbls.dLblPos = 'outEnd'
    dLbls.numFmt = '0.0%'
    individual_label = MagicMock()
    individual_label.idx = 0
    individual_label.showVal = False
    dLbls.dLbl = [individual_label]
    with patch.object(extractor, '_extract_individual_data_label', return_value={'index': 0, 'show_value': False}):
        labels_info = extractor.extract_data_labels(series)
    assert labels_info['position'] == 'outEnd'
    assert labels_info['number_format'] == '0.0%'
    assert len(labels_info['labels']) == 1
    assert labels_info['labels'][0]['show_value'] is False

def test_extract_legend_entry(extractor):
    """测试 _extract_legend_entry 辅助函数。"""
    entry = MagicMock()
    entry.idx = 1
    entry.delete = False
    run = MagicMock()
    type(run).t = PropertyMock(return_value='Legend Text')
    p = MagicMock()
    p.r = [run]
    type(entry.txPr).p = PropertyMock(return_value=[p])
    entry_info = extractor._extract_legend_entry(entry)
    assert entry_info['index'] == 1
    assert entry_info['delete'] is False
    assert entry_info['text'] == 'Legend Text'

def test_extract_textbox_annotations_from_plotarea(extractor):
    """测试从 plotArea 提取文本框注释。"""
    chart = MagicMock()
    # 模拟 plotArea.txPr
    run = MagicMock()
    type(run).t = PropertyMock(return_value='Plot Area Text')
    p = MagicMock()
    p.r = [run]
    type(chart.plotArea.txPr).p = PropertyMock(return_value=[p])
    
    # 显式地将其他检查的属性设置为None，以防止MagicMock创建它们
    chart.plotArea.tx = None
    chart.plotArea.text = None
    chart.plotArea.textBox = None
    chart.plotArea.txBox = None
    chart.textBox = None
    chart.txBox = None
    chart.freeText = None
    chart.annotation = None
    
    annotations = extractor._extract_textbox_annotations(chart)
    assert len(annotations) == 1
    assert annotations[0]['text'] == 'Plot Area Text'
    assert annotations[0]['source'] == 'plotArea.txPr'

def test_extract_shape_annotations_from_attribute(extractor):
    """测试从 chart.shapes 提取形状注释。"""
    chart = MagicMock()
    
    # 模拟 shape 对象
    shape1 = MagicMock()
    chart.shapes = [shape1]
    
    # 显式地将其他检查的属性设置为None
    chart.plotArea.sp = None
    chart.plotArea.shape = None
    chart.plotArea.shapes = None
    chart.plotArea.spPr = None
    chart.sp = None
    chart.drawing = None
    chart.drawingElements = None
    chart.graphicalProperties = None
    
    # 模拟 _extract_single_shape 的返回值
    mock_shape_info = {'text': 'Shape 1 Text', 'shape_type': 'rect'}
    
    with patch.object(extractor, '_extract_single_shape', return_value=mock_shape_info) as mock_extract:
        annotations = extractor._extract_shape_annotations(chart)
    
    # 验证 _extract_single_shape 被正确调用
    mock_extract.assert_called_once_with(shape1, 'chart.shapes[0]')
    
    assert len(annotations) == 1
    assert annotations[0] == mock_shape_info