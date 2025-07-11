import pytest
from src.models.table_model import Sheet, Row, Cell, Style
from src.converters.html_converter import HTMLConverter
from src.converters.json_converter import JSONConverter

def test_html_converter():
    # 1. 准备：创建一个简单的 Sheet 对象
    sheet = Sheet(
        name="Test Sheet",
        rows=[
            Row(cells=[Cell(value='A1'), Cell(value='B1')]),
            Row(cells=[Cell(value='A2'), Cell(value='B2')])
        ]
    )

    # 2. 执行：将 Sheet 转换为 HTML 字符串
    converter = HTMLConverter()
    html_output = converter.convert(sheet)

    # 3. 断言：检查输出是否为有效的 HTML 表格
    assert '<table>' in html_output
    assert '</table>' in html_output

    # 检查基本表格结构（兼容优化输出）
    assert 'A1' in html_output
    assert 'B1' in html_output
    assert 'A2' in html_output
    assert 'B2' in html_output
    # 注意：优化后的HTML可能不包含默认的colspan="1"和rowspan="1"
    # 只检查基本的表格结构


def test_json_converter_basic_conversion():
    """测试基本的 JSON 转换功能。"""
    # 创建一个简单的 sheet
    sheet = Sheet(
        name="TestSheet",
        rows=[
            Row(cells=[Cell(value="Header1"), Cell(value="Header2")]),
            Row(cells=[Cell(value="Data1"), Cell(value="Data2")])
        ]
    )

    converter = JSONConverter()
    json_data = converter.convert(sheet)

    # 测试元数据
    assert json_data['metadata']['name'] == "TestSheet"
    assert json_data['metadata']['rows'] == 2
    assert json_data['metadata']['cols'] == 2
    assert json_data['metadata']['has_merged_cells'] == False

    # 测试数据结构
    assert len(json_data['data']) == 2
    assert json_data['data'][0]['row'] == 0
    assert len(json_data['data'][0]['cells']) == 2
    assert json_data['data'][0]['cells'][0]['value'] == "Header1"
    assert json_data['data'][0]['cells'][0]['col'] == 0


def test_json_converter_with_styles():
    """测试带样式单元格的 JSON 转换。"""
    # 创建一个带样式单元格的 sheet
    styled_cell = Cell(value="Styled", style=Style(bold=True, font_color="#FF0000"))
    normal_cell = Cell(value="Normal")

    sheet = Sheet(
        name="StyledSheet",
        rows=[Row(cells=[styled_cell, normal_cell])]
    )

    converter = JSONConverter()
    json_data = converter.convert(sheet)

    # 测试样式提取
    assert len(json_data['styles']) > 0

    # 测试样式引用
    first_cell = json_data['data'][0]['cells'][0]
    second_cell = json_data['data'][0]['cells'][1]

    assert first_cell['style_id'] is not None
    assert second_cell['style_id'] is None

    # 测试样式内容
    style_id = first_cell['style_id']
    style_data = json_data['styles'][style_id]
    assert style_data['bold'] == True
    assert style_data['font_color'] == "#FF0000"


def test_json_converter_with_merged_cells():
    """测试带合并单元格的 JSON 转换。"""
    # 创建一个带合并单元格的 sheet
    sheet = Sheet(
        name="MergedSheet",
        rows=[Row(cells=[Cell(value="Merged"), Cell(value="Cell2")])],
        merged_cells=["A1:B1"]
    )

    converter = JSONConverter()
    json_data = converter.convert(sheet)

    # 测试合并单元格元数据
    assert json_data['metadata']['has_merged_cells'] == True
    assert json_data['metadata']['merged_cells_count'] == 1
    assert json_data['merged_cells'] == ["A1:B1"]


def test_json_converter_empty_sheet():
    """测试空 sheet 的 JSON 转换。"""
    # 创建一个空的 sheet
    sheet = Sheet(name="EmptySheet", rows=[])

    converter = JSONConverter()
    json_data = converter.convert(sheet)

    # 测试空 sheet 处理
    assert json_data['metadata']['rows'] == 0
    assert json_data['metadata']['cols'] == 0
    assert len(json_data['data']) == 0
    assert len(json_data['styles']) == 0


def test_json_converter_to_string():
    """Test JSON string conversion."""
    # Create a simple sheet
    sheet = Sheet(
        name="StringTest",
        rows=[Row(cells=[Cell(value="Test")])]
    )

    converter = JSONConverter()

    # Test compact JSON string
    json_string = converter.to_json_string(sheet)
    assert isinstance(json_string, str)
    assert "StringTest" in json_string

    # Test formatted JSON string
    formatted_json = converter.to_json_string(sheet, indent=2)
    assert "\n" in formatted_json  # Should have newlines when formatted


def test_json_converter_size_estimation():
    """Test JSON size estimation functionality."""
    # Create a sheet with multiple rows and styles
    sheet = Sheet(
        name="SizeTest",
        rows=[
            Row(cells=[Cell(value="Header", style=Style(bold=True))]),
            Row(cells=[Cell(value="Data1")]),
            Row(cells=[Cell(value="Data2")])
        ]
    )

    converter = JSONConverter()
    size_info = converter.estimate_json_size(sheet)

    # Test size estimation structure
    assert 'total_characters' in size_info
    assert 'total_bytes' in size_info
    assert 'rows' in size_info
    assert 'unique_styles' in size_info
    assert 'cells' in size_info

    # Test size values
    assert size_info['rows'] == 3
    assert size_info['cells'] == 3
    assert size_info['total_characters'] > 0
    assert size_info['total_bytes'] > 0


def test_json_converter_none_values():
    """Test JSON conversion with None values."""
    # Create a sheet with None values
    sheet = Sheet(
        name="NoneTest",
        rows=[Row(cells=[Cell(value=None), Cell(value="NotNone")])]
    )

    converter = JSONConverter()
    json_data = converter.convert(sheet)

    # Test None value handling
    first_cell = json_data['data'][0]['cells'][0]
    second_cell = json_data['data'][0]['cells'][1]

    assert first_cell['value'] is None
    assert second_cell['value'] == "NotNone"


def test_json_converter_error_handling():
    """Test JSON converter error handling."""
    converter = JSONConverter()

    # Test None sheet
    with pytest.raises(ValueError):
        converter.convert(None)


def test_html_converter_optimization():
    """Test HTML converter optimization features."""
    # Create a sheet with multiple styled cells
    sheet = Sheet(
        name="OptimizationTest",
        rows=[
            Row(cells=[
                Cell(value="Header1", style=Style(bold=True, font_color="#FF0000")),
                Cell(value="Header2", style=Style(bold=True, font_color="#FF0000")),
                Cell(value="Header3", style=Style(italic=True, font_color="#0000FF"))
            ]),
            Row(cells=[
                Cell(value="Data1", style=Style(bold=True, font_color="#FF0000")),
                Cell(value="Data2"),
                Cell(value="Data3", style=Style(italic=True, font_color="#0000FF"))
            ])
        ]
    )

    converter = HTMLConverter(compact_mode=True)

    # Test optimized conversion
    optimized_html = converter.convert(sheet, optimize=True)
    assert isinstance(optimized_html, str)
    assert len(optimized_html) > 0
    assert "class=" in optimized_html  # Should contain CSS classes

    # Test that CSS classes are generated
    css_classes = converter._generate_css_classes(sheet)
    assert len(css_classes) > 0  # Should have unique styles

    # Test CSS generation
    css_styles = converter.generate_css_styles(css_classes)
    assert isinstance(css_styles, str)
    assert "font-weight: bold" in css_styles or "color: #FF0000" in css_styles


def test_html_converter_css_class_reuse():
    """Test CSS class reuse for identical styles."""
    # Create cells with identical styles
    same_style = Style(bold=True, font_color="#FF0000")
    sheet = Sheet(
        name="ClassReuseTest",
        rows=[
            Row(cells=[
                Cell(value="Cell1", style=same_style),
                Cell(value="Cell2", style=same_style),
                Cell(value="Cell3", style=Style(italic=True))  # Different style
            ])
        ]
    )

    converter = HTMLConverter()
    css_classes = converter._generate_css_classes(sheet)

    # Should have only 2 unique styles (same_style and italic style)
    assert len(css_classes) == 2

    # Test that cells with same style get same CSS class
    cell1_class = converter.get_cell_css_class(sheet.rows[0].cells[0], css_classes)
    cell2_class = converter.get_cell_css_class(sheet.rows[0].cells[1], css_classes)
    cell3_class = converter.get_cell_css_class(sheet.rows[0].cells[2], css_classes)

    assert cell1_class == cell2_class  # Same style, same class
    assert cell1_class != cell3_class  # Different style, different class


def test_html_converter_size_reduction():
    """Test HTML size reduction estimation."""
    # Create a sheet with repeated styles
    repeated_style = Style(bold=True, font_color="#FF0000", background_color="#FFFF00")
    sheet = Sheet(
        name="SizeTest",
        rows=[
            Row(cells=[Cell(value=f"Cell{i}", style=repeated_style) for i in range(5)])
            for _ in range(3)
        ]
    )

    converter = HTMLConverter(compact_mode=True)
    size_info = converter.estimate_size_reduction(sheet)

    # Test size reduction structure
    assert 'original_size' in size_info
    assert 'optimized_size' in size_info
    assert 'size_reduction' in size_info
    assert 'reduction_percentage' in size_info

    # Test that optimization actually reduces size
    assert size_info['optimized_size'] < size_info['original_size']
    assert size_info['size_reduction'] > 0
    assert size_info['reduction_percentage'] > 0


def test_html_converter_compression():
    """Test HTML compression functionality."""
    # Create a simple sheet
    sheet = Sheet(
        name="CompressionTest",
        rows=[Row(cells=[Cell(value="Test")])]
    )

    converter_compact = HTMLConverter(compact_mode=True)
    converter_normal = HTMLConverter(compact_mode=False)

    # Generate HTML with and without compression
    compact_html = converter_compact.convert(sheet, optimize=False)
    normal_html = converter_normal.convert(sheet, optimize=False)

    # Compact HTML should be smaller (less whitespace)
    assert len(compact_html) <= len(normal_html)


def test_html_converter_css_generation():
    """测试 CSS 样式生成。"""
    converter = HTMLConverter()

    # 创建测试样式
    css_classes = {
        'test1': Style(bold=True, font_color="#FF0000"),
        'test2': Style(italic=True, background_color="#FFFF00"),
        'test3': Style(font_size=14.0, font_name="Arial")
    }

    css_styles = converter.generate_css_styles(css_classes)

    # 测试 CSS 内容
    assert isinstance(css_styles, str)
    assert ".test1" in css_styles
    assert "font-weight: bold" in css_styles
    assert "color: #FF0000" in css_styles
    assert ".test2" in css_styles
    assert "font-style: italic" in css_styles
    assert "background-color: #FFFF00" in css_styles


def test_html_converter_empty_sheet_optimization():
    """测试空 sheet 的优化。"""
    empty_sheet = Sheet(name="EmptyTest", rows=[])

    converter = HTMLConverter()

    # 输出中应包含 sheet 名称
    css_classes = converter._generate_css_classes(empty_sheet)
    assert len(css_classes) == 0

    optimized_html = converter.convert(empty_sheet, optimize=True)
    assert isinstance(optimized_html, str)
    assert "EmptyTest" in optimized_html


def test_html_converter_error_handling():
    """Test HTML converter error handling."""
    converter = HTMLConverter()

    # 测试空的sheet
    with pytest.raises(ValueError):
        converter.convert(None)
