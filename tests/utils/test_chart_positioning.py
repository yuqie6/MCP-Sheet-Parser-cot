import pytest
from unittest.mock import MagicMock

from src.utils.chart_positioning import ChartPositionCalculator, ChartCSSPosition
from src.models.table_model import Sheet, ChartPosition

@pytest.fixture
def mock_sheet():
    """提供一个模拟的 Sheet 对象，用于测试。"""
    sheet = MagicMock(spec=Sheet)
    sheet.column_widths = {0: 10, 1: 12, 2: 8}
    sheet.row_heights = {0: 20, 1: 25, 2: 18}
    sheet.default_column_width = 9
    sheet.default_row_height = 15
    return sheet

@pytest.fixture
def calculator(mock_sheet):
    """提供一个 ChartPositionCalculator 实例。"""
    return ChartPositionCalculator(mock_sheet)

def test_initialization(calculator, mock_sheet):
    """测试计算器是否正确初始化。"""
    assert calculator.sheet == mock_sheet

def test_calculate_cell_position_single_cell(calculator):
    """测试计算单个单元格的位置。"""
    # A1单元格 (0, 0)
    x, y = calculator._calculate_cell_position(0, 0)
    assert x == 0
    assert y == 0

    # B2单元格 (1, 1)
    x, y = calculator._calculate_cell_position(1, 1)
    expected_x = 10 * 8.45  # col 0 width
    expected_y = 20  # row 0 height
    assert abs(x - expected_x) < 1e-9
    assert abs(y - expected_y) < 1e-9

@pytest.fixture
def mock_chart_position():
    """提供一个模拟的 ChartPosition 对象。"""
    pos = MagicMock(spec=ChartPosition)
    pos.from_row = 1
    pos.from_col = 1
    pos.from_row_offset = 100000
    pos.from_col_offset = 50000
    pos.to_row = 3
    pos.to_col = 3
    pos.to_row_offset = 200000
    pos.to_col_offset = 150000
    return pos

def test_calculate_chart_width(calculator, mock_chart_position):
    """测试计算图表的宽度（修正版）。"""
    width = calculator._calculate_chart_width(mock_chart_position)

    # 根据源代码的逻辑重新计算期望值
    start_offset_px = mock_chart_position.from_col_offset * calculator.EMU_TO_PX
    end_offset_px = mock_chart_position.to_col_offset * calculator.EMU_TO_PX

    from_col_width_px = calculator.sheet.column_widths.get(mock_chart_position.from_col, calculator.sheet.default_column_width) * calculator.EXCEL_TO_PX
    
    middle_cols_width_px = 0
    for i in range(mock_chart_position.from_col + 1, mock_chart_position.to_col):
        middle_cols_width_px += calculator.sheet.column_widths.get(i, calculator.sheet.default_column_width) * calculator.EXCEL_TO_PX

    expected_width = (from_col_width_px - start_offset_px) + middle_cols_width_px + end_offset_px
    
    # 源代码中有 max(50, width) 的逻辑
    assert abs(width - max(50, expected_width)) < 1e-9

def test_calculate_chart_height(calculator, mock_chart_position):
    """测试计算图表的高度（修正版）。"""
    height = calculator._calculate_chart_height(mock_chart_position)

    # 根据源代码的逻辑重新计算期望值
    start_offset_pt = mock_chart_position.from_row_offset * calculator.EMU_TO_PT
    end_offset_pt = mock_chart_position.to_row_offset * calculator.EMU_TO_PT

    from_row_height_pt = calculator.sheet.row_heights.get(mock_chart_position.from_row, calculator.sheet.default_row_height)
    
    middle_rows_height_pt = 0
    for i in range(mock_chart_position.from_row + 1, mock_chart_position.to_row):
        middle_rows_height_pt += calculator.sheet.row_heights.get(i, calculator.sheet.default_row_height)

    expected_height = (from_row_height_pt - start_offset_pt) + middle_rows_height_pt + end_offset_pt
    
    # 源代码中有 max(50, height) 的逻辑
    assert abs(height - max(50, expected_height)) < 1e-9

def test_calculate_chart_css_position(calculator, mock_chart_position):
    """测试获取最终的CSS位置对象（修正版）。"""
    # 修正方法调用
    css_pos = calculator.calculate_chart_css_position(mock_chart_position)

    assert isinstance(css_pos, ChartCSSPosition)
    
    # 验证 top 和 left 是否正确计算
    start_x_base, start_y_base = calculator._calculate_cell_position(mock_chart_position.from_col, mock_chart_position.from_row)
    
    col_offset_px = mock_chart_position.from_col_offset * calculator.EMU_TO_PX
    row_offset_pt = mock_chart_position.from_row_offset * calculator.EMU_TO_PT

    expected_left = start_x_base + col_offset_px
    expected_top = start_y_base + row_offset_pt

    assert abs(css_pos.left - expected_left) < 1e-9
    assert abs(css_pos.top - expected_top) < 1e-9
    
    # 验证 width 和 height 是否与独立方法计算结果一致
    expected_width = calculator._calculate_chart_width(mock_chart_position)
    expected_height = calculator._calculate_chart_height(mock_chart_position)
    
    assert abs(css_pos.width - expected_width) < 1e-9
    assert abs(css_pos.height - expected_height) < 1e-9

def test_calculate_cell_position_multi_cell(calculator):
    """测试计算跨多个单元格的位置。"""
    # C3单元格 (2, 2)
    x, y = calculator._calculate_cell_position(2, 2)
    expected_x = (10 + 12) * 8.45  # col 0 + col 1 width
    expected_y = 20 + 25  # row 0 + row 1 height
    assert abs(x - expected_x) < 1e-9
    assert abs(y - expected_y) < 1e-9

@pytest.fixture
def mock_single_cell_chart_position():
    """提供一个在单个单元格内的模拟 ChartPosition 对象。"""
    pos = MagicMock(spec=ChartPosition)
    pos.from_row = 1
    pos.from_col = 1
    pos.from_row_offset = 100000  # ~1/9th of the way in
    pos.from_col_offset = 50000   # ~1/18th of the way in
    pos.to_row = 1 # Same row
    pos.to_col = 1 # Same col
    pos.to_row_offset = 800000
    pos.to_col_offset = 850000
    return pos

def test_calculate_chart_width_single_cell(calculator, mock_single_cell_chart_position):
    """测试在单个单元格内计算图表宽度。"""
    width = calculator._calculate_chart_width(mock_single_cell_chart_position)
    
    start_offset_px = mock_single_cell_chart_position.from_col_offset * calculator.EMU_TO_PX
    end_offset_px = mock_single_cell_chart_position.to_col_offset * calculator.EMU_TO_PX
    
    expected_width = end_offset_px - start_offset_px
    assert abs(width - max(50, expected_width)) < 1e-9

def test_calculate_chart_height_single_row(calculator, mock_single_cell_chart_position):
    """测试在单行内计算图表高度。"""
    height = calculator._calculate_chart_height(mock_single_cell_chart_position)

    start_offset_pt = mock_single_cell_chart_position.from_row_offset * calculator.EMU_TO_PT
    end_offset_pt = mock_single_cell_chart_position.to_row_offset * calculator.EMU_TO_PT

    expected_height = end_offset_pt - start_offset_pt
    assert abs(height - max(50, expected_height)) < 1e-9

def test_calculate_image_dimensions(calculator, mock_chart_position):
    """测试图片尺寸的计算。"""
    # Test width
    width = calculator._calculate_image_width(mock_chart_position)
    width_emu = mock_chart_position.to_col_offset - mock_chart_position.from_col_offset
    expected_width_px = width_emu * calculator.EMU_TO_PX
    assert abs(width - max(20, expected_width_px)) < 1e-9

    # Test height
    height = calculator._calculate_image_height(mock_chart_position)
    height_emu = mock_chart_position.to_row_offset - mock_chart_position.from_row_offset
    expected_height_pt = height_emu * calculator.EMU_TO_PT
    assert abs(height - max(15, expected_height_pt)) < 1e-9

def test_generate_chart_html_with_positioning_for_image(calculator, mock_single_cell_chart_position):
    """测试为图片生成带定位的HTML。"""
    mock_chart = MagicMock()
    mock_chart.type = 'image'
    # 修正：通过参数注入 Fixture，而不是直接调用
    mock_chart.position = mock_single_cell_chart_position
    
    # Mock the _calculate_image_position to return a predictable ChartCSSPosition
    css_pos = ChartCSSPosition(left=10, top=20, width=100, height=80)
    calculator._calculate_image_position = MagicMock(return_value=css_pos)

    chart_html = "<svg>...</svg>"
    positioned_html = calculator.generate_chart_html_with_positioning(mock_chart, chart_html)
    
    calculator._calculate_image_position.assert_called_once_with(mock_chart.position)
    assert 'left: 10.0px;' in positioned_html
    assert 'top: 26.7px;' in positioned_html # 20pt * 1.333
    assert 'width: 100.0px;' in positioned_html
    assert 'height: 106.6px;' in positioned_html # 80pt * 1.333
    assert chart_html in positioned_html

def test_get_chart_overlay_css(calculator, mock_chart_position):
    """测试生成图表覆盖的CSS。"""
    css_string = calculator.get_chart_overlay_css(mock_chart_position, "my-container")
    
    css_pos = calculator.calculate_chart_css_position(mock_chart_position)

    assert '#my-container' in css_string
    assert f"left: {css_pos.left:.1f}px;" in css_string
    assert f"top: {css_pos.top:.1f}pt;" in css_string
    assert f"width: {css_pos.width:.1f}px;" in css_string
    assert f"height: {css_pos.height:.1f}pt;" in css_string

def test_calculate_cell_position_with_defaults(calculator):
    """测试计算位置时使用默认行高和列宽。"""
    # E5单元格 (4, 4)
    x, y = calculator._calculate_cell_position(4, 4)
    expected_x = (10 + 12 + 8 + 9) * 8.45 # col 0, 1, 2, 3 (default)
    expected_y = 20 + 25 + 18 + 15 # row 0, 1, 2, 3 (default)
    assert abs(x - expected_x) < 1e-9
    assert abs(y - expected_y) < 1e-9