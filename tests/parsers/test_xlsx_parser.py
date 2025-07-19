import pytest
from unittest.mock import MagicMock, patch, mock_open, PropertyMock, call
from src.parsers.xlsx_parser import XlsxParser, XlsxRowProvider
from src.models.table_model import Sheet, LazySheet, Chart, Row, Cell
import openpyxl
# 明确导入 openpyxl 的内部类型，以解决 Pylance 的类型检查错误
from openpyxl.cell.cell import Cell as OpenpyxlCell
from openpyxl.worksheet.worksheet import Worksheet as OpenpyxlWorksheet
from openpyxl.chart.bar_chart import BarChart as OpenpyxlBarChart
from openpyxl.drawing.image import Image as OpenpyxlImage
import zipfile
from io import BytesIO

@pytest.fixture
def mock_openpyxl_cell():
    """提供一个模拟的 openpyxl 单元格对象。"""
    cell = MagicMock(spec=OpenpyxlCell)
    cell.value = "Test Data"
    cell.data_type = 's'
    cell.font = MagicMock()
    cell.fill = MagicMock()
    cell.border = MagicMock()
    cell.alignment = MagicMock()
    cell.hyperlink = None
    cell.comment = None
    cell.has_style = True
    return cell

@pytest.fixture
def mock_openpyxl_worksheet(mock_openpyxl_cell):
    """提供一个模拟的 openpyxl 工作表对象。"""
    worksheet = MagicMock(spec=OpenpyxlWorksheet)
    worksheet.title = "TestSheet"
    worksheet.max_row = 2
    worksheet.max_column = 2
    worksheet.cell.return_value = mock_openpyxl_cell
    # 模拟 openpyxl 的 merged_cells 对象，它有一个 .ranges 属性
    mock_merged_cells = MagicMock()
    mock_merged_cells.ranges = ["A1:B1"]
    worksheet.merged_cells = mock_merged_cells
    
    col_dim = MagicMock()
    col_dim.width = 10
    worksheet.column_dimensions = {'A': col_dim, 'B': col_dim}
    row_dim = MagicMock()
    row_dim.height = 20
    worksheet.row_dimensions = {1: row_dim, 2: row_dim}

    # Fix: Mock sheet_format attribute
    sheet_format = MagicMock()
    sheet_format.defaultColWidth = 8.43
    sheet_format.defaultRowHeight = 18.0
    worksheet.sheet_format = sheet_format
    
    worksheet._charts = []
    worksheet._images = []
    return worksheet

@pytest.fixture
def mock_openpyxl_workbook(mock_openpyxl_worksheet):
    """Fixture for a mocked openpyxl workbook."""
    workbook = MagicMock(spec=openpyxl.Workbook)
    workbook.sheetnames = ["TestSheet"]
    workbook.__getitem__.return_value = mock_openpyxl_worksheet
    type(workbook).active = PropertyMock(return_value=mock_openpyxl_worksheet)
    workbook.close = MagicMock()
    return workbook

class TestXlsxRowProvider:
    @patch('openpyxl.load_workbook')
    def test_get_total_rows(self, mock_load_workbook, mock_openpyxl_workbook):
        mock_load_workbook.return_value = mock_openpyxl_workbook
        provider = XlsxRowProvider("dummy.xlsx")
        total_rows = provider.get_total_rows()
        assert total_rows == 2
        # Test caching
        assert provider._total_rows_cache == 2
        provider.get_total_rows()
        mock_load_workbook.assert_called_once()

    @patch('openpyxl.load_workbook')
    def test_get_merged_cells(self, mock_load_workbook, mock_openpyxl_workbook):
        mock_load_workbook.return_value = mock_openpyxl_workbook
        provider = XlsxRowProvider("dummy.xlsx")
        merged_cells = provider._get_merged_cells()
        assert merged_cells == ["A1:B1"]
        # Test caching
        assert provider._merged_cells_cache == ["A1:B1"]
        provider._get_merged_cells()
        mock_load_workbook.assert_called_once()

    @patch('openpyxl.load_workbook')
    def test_get_row(self, mock_load_workbook, mock_openpyxl_workbook):
        mock_load_workbook.return_value = mock_openpyxl_workbook
        provider = XlsxRowProvider("dummy.xlsx")
        row = provider.get_row(0)
        assert isinstance(row, Row)
        assert len(row.cells) == 2
        assert row.cells[0].value == "Test Data"

    @patch('openpyxl.load_workbook')
    def test_get_row_out_of_range(self, mock_load_workbook, mock_openpyxl_workbook):
        mock_load_workbook.return_value = mock_openpyxl_workbook
        provider = XlsxRowProvider("dummy.xlsx")
        with pytest.raises(IndexError):
            provider.get_row(5)

    def test_provider_load_failure(self):
        """测试当底层openpyxl加载失败时，提供者是否能正确抛出异常。"""
        provider = XlsxRowProvider("dummy.xlsx")
        with patch('openpyxl.load_workbook', side_effect=Exception("Failed to load")):
            with pytest.raises(RuntimeError, match="流式读取XLSX文件失败"):
                list(provider.iter_rows())
        
        with patch('openpyxl.load_workbook', side_effect=Exception("Failed to load")):
            with pytest.raises(RuntimeError, match="无法加载工作簿以获取总行数"):
                provider.get_total_rows()
        
        with patch('openpyxl.load_workbook', side_effect=Exception("Failed to load")):
            with pytest.raises(RuntimeError, match="无法加载工作簿以获取合并单元格"):
                provider._get_merged_cells()

    def test_get_worksheet_info_exception_handling(self):
        """测试获取工作表信息时的异常处理。"""
        provider = XlsxRowProvider("nonexistent.xlsx")

        # 测试异常情况下的处理
        worksheet_title = provider._get_worksheet_info()
        assert worksheet_title == ""

        # 测试缓存机制
        worksheet_title2 = provider._get_worksheet_info()
        assert worksheet_title2 == ""

    def test_get_worksheet_info_with_sheet_name(self, tmp_path):
        """测试指定工作表名称获取信息。"""
        # 创建一个真实的XLSX文件
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet1 = workbook.active
        sheet1.title = "Sheet1"
        sheet2 = workbook.create_sheet("Sheet2")
        workbook.save(file_path)

        # 测试指定工作表名称
        provider = XlsxRowProvider(str(file_path), "Sheet2")
        worksheet_title = provider._get_worksheet_info()
        assert worksheet_title == "Sheet2"

    def test_get_worksheet_info_none_worksheet(self):
        """测试工作表为None的情况。"""
        provider = XlsxRowProvider("test.xlsx")

        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_workbook.active = None
            mock_load.return_value = mock_workbook

            worksheet_title = provider._get_worksheet_info()
            assert worksheet_title == ""

    def test_iter_rows_with_parameters(self, tmp_path):
        """测试带参数的行迭代。"""
        # 创建一个真实的XLSX文件
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for i in range(5):
            for j in range(3):
                sheet.cell(row=i+1, column=j+1, value=f"Cell{i}_{j}")
        workbook.save(file_path)

        provider = XlsxRowProvider(str(file_path))

        # 测试从第2行开始，最多2行
        rows = list(provider.iter_rows(start_row=1, max_rows=2))
        assert len(rows) == 2
        assert rows[0].cells[0].value == "Cell1_0"

    def test_iter_rows_none_worksheet(self):
        """测试工作表为None时的行迭代。"""
        provider = XlsxRowProvider("test.xlsx")

        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_workbook.active = None
            mock_load.return_value = mock_workbook

            rows = list(provider.iter_rows())
            assert len(rows) == 0

    def test_iter_rows_exception_handling(self):
        """测试行迭代时的异常处理。"""
        provider = XlsxRowProvider("nonexistent.xlsx")

        # 测试文件不存在时的异常处理
        with pytest.raises(RuntimeError, match="流式读取XLSX文件失败"):
            list(provider.iter_rows())

class TestXlsxParser:
    @patch('openpyxl.load_workbook')
    def test_xlsx_parser_parse(self, mock_load_workbook, mock_openpyxl_workbook):
        """Test basic parsing of an XLSX file."""
        mock_load_workbook.return_value = mock_openpyxl_workbook
        parser = XlsxParser()
        sheets = parser.parse("dummy.xlsx")
        assert len(sheets) == 1
        assert isinstance(sheets[0], Sheet)
        assert sheets[0].name == "TestSheet"
        assert sheets[0].rows[0].cells[0].value == "Test Data"
        assert sheets[0].merged_cells == ["A1:B1"]

    @patch('openpyxl.load_workbook')
    def test_formula_parsing(self, mock_load_workbook, mock_openpyxl_workbook, mock_openpyxl_worksheet):
        """测试公式单元格的解析，确保公式和计算结果都被正确提取。"""
        formula_cell = MagicMock(spec=OpenpyxlCell)
        formula_cell.value = "=SUM(A1:B1)"
        formula_cell.data_type = 'f'
        # Mock style attributes
        formula_cell.font = MagicMock()
        formula_cell.fill = MagicMock()
        formula_cell.border = MagicMock()
        formula_cell.alignment = MagicMock()
        formula_cell.hyperlink = None
        formula_cell.comment = None
        formula_cell.has_style = True

        data_only_cell = MagicMock()
        data_only_cell.value = 42

        mock_openpyxl_worksheet.cell.side_effect = [formula_cell, data_only_cell] * 2 # for both workbooks
        
        data_only_workbook = MagicMock()
        data_only_worksheet = MagicMock()
        data_only_worksheet.cell.return_value = data_only_cell
        data_only_workbook.__getitem__.return_value = data_only_worksheet
        
        mock_load_workbook.side_effect = [mock_openpyxl_workbook, data_only_workbook]
        
        parser = XlsxParser()
        sheets = parser.parse("dummy.xlsx")
        
        parsed_cell = sheets[0].rows[0].cells[0]
        assert parsed_cell.value == 42
        assert parsed_cell.formula == "=SUM(A1:B1)"

    @patch('src.parsers.xlsx_parser.XlsxParser._fix_excel_styles', return_value="fixed.xlsx")
    def test_load_failure_and_fix(self, mock_fix):
        """测试当加载因样式问题失败时，是否会触发修复机制。"""
        parser = XlsxParser()
        final_workbook = MagicMock()

        # 创建一个更智能的 side_effect 函数
        def smart_side_effect(file_path, **kwargs):
            if file_path == "dummy.xlsx":
                # 首次加载原始文件时，抛出特定异常以触发修复
                raise zipfile.BadZipFile("File is not a zip file or contains invalid 'Fill'")
            elif file_path == "fixed.xlsx":
                # 加载修复后的文件时，返回成功的 workbook 对象
                return final_workbook
            else:
                # 其他任何情况都抛出通用异常
                raise Exception("Unexpected file load attempt")

        with patch('openpyxl.load_workbook', side_effect=smart_side_effect) as mock_load_workbook:
            with patch('logging.getLogger'):
                with patch('src.parsers.xls_parser.XlsParser.parse', return_value=[]):
                    parser.parse("dummy.xlsx")
                    # 验证修复函数被调用
                    mock_fix.assert_called_once_with("dummy.xlsx")
                    # 验证代码尝试加载了修复后的文件
                    mock_load_workbook.assert_any_call("fixed.xlsx", data_only=False, keep_vba=False, keep_links=False)

    @patch('openpyxl.load_workbook', side_effect=Exception("All loads fail"))
    @patch('src.parsers.xls_parser.XlsParser.parse', return_value=[Sheet(name="xls_sheet", rows=[])])
    def test_fallback_to_xls_parser(self, mock_xls_parse, mock_load_workbook):
        parser = XlsxParser()
        with patch('src.parsers.xls_parser.XlsParser') as mock_xls_parser_class:
            mock_xls_parser_instance = mock_xls_parser_class.return_value
            mock_xls_parser_instance.parse.return_value = [Sheet(name="xls_sheet", rows=[])]
            
            sheets = parser.parse("dummy.xlsx")
            
            mock_xls_parser_instance.parse.assert_called_once_with("dummy.xlsx")
            assert sheets[0].name == "xls_sheet"

    def test_xlsx_parser_streaming_support(self):
        """Test streaming support declaration."""
        parser = XlsxParser()
        assert parser.supports_streaming() is True

    @patch('src.parsers.xlsx_parser.XlsxRowProvider')
    def test_create_lazy_sheet(self, MockXlsxRowProvider):
        """Test create_lazy_sheet for XLSX."""
        mock_provider_instance = MockXlsxRowProvider.return_value
        mock_provider_instance._get_worksheet_info.return_value = "LazySheet"
        mock_provider_instance._get_merged_cells.return_value = []

        parser = XlsxParser()
        lazy_sheet = parser.create_lazy_sheet("dummy.xlsx")

        assert isinstance(lazy_sheet, LazySheet)
        assert lazy_sheet.name == "LazySheet"
        MockXlsxRowProvider.assert_called_once_with("dummy.xlsx", None)

    @patch('openpyxl.load_workbook')
    def test_extract_charts(self, mock_load_workbook, mock_openpyxl_workbook, mock_openpyxl_worksheet):
        """测试图表提取功能，确保图表被正确解析并添加到Sheet对象中。"""
        chart = MagicMock(spec=OpenpyxlBarChart)
        chart.title = "My Chart"
        chart.anchor = None
        mock_openpyxl_worksheet._charts = [chart]
        mock_load_workbook.return_value = mock_openpyxl_workbook
        parser = XlsxParser()
        with patch.object(parser.chart_extractor, 'extract_axis_title', return_value="My Chart"), \
             patch.object(parser, '_extract_chart_data', return_value={'type': 'bar'}):
            sheets = parser.parse("dummy.xlsx")
            # 修复断言，确保在有图表时，charts列表不为空
            assert len(sheets[0].charts) > 0, "图表列表不应为空"
            assert sheets[0].charts[0].name == "My Chart"

    @patch('openpyxl.load_workbook')
    def test_extract_images(self, mock_load_workbook, mock_openpyxl_workbook, mock_openpyxl_worksheet):
        """测试图片提取功能，确保图片被解析为特殊的图表对象。"""
        image = MagicMock(spec=OpenpyxlImage)
        image.ref = BytesIO(b"imagedata")
        image.anchor._from.col = 1
        image.anchor._from.row = 1
        image.anchor._from.colOff = 0
        image.anchor._from.rowOff = 0
        image.anchor.to.col = 3
        image.anchor.to.row = 5
        image.anchor.to.colOff = 0
        image.anchor.to.rowOff = 0
        
        mock_openpyxl_worksheet._images = [image]
        mock_load_workbook.return_value = mock_openpyxl_workbook
        
        parser = XlsxParser()
        sheets = parser.parse("dummy.xlsx")
        
        assert len(sheets[0].charts) == 1, "应从工作表中提取一个图片作为图表对象"
        chart = sheets[0].charts[0]
        assert chart.type == "image", "图表类型应为 'image'"
        
        # 增加健壮性检查，确保 chart_data 和 position 不是 None
        assert chart.chart_data is not None, "chart_data 不应为 None"
        assert chart.position is not None, "position 不应为 None"
        
        assert chart.chart_data.get('image_data') == b"imagedata", "图片数据不匹配"
        assert chart.position.from_col == 1, "起始列不正确"
        assert chart.position.from_row == 1, "起始行不正确"

    def test_fix_excel_styles(self):
        """Test the style fixing logic."""
        parser = XlsxParser()
        
        # Create a mock zip file with a broken styles.xml
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zf:
            zf.writestr("xl/styles.xml", '<styleSheet><fills count="2"><fill/><fill><patternFill patternType="none"/></fill></fills></styleSheet>')
            zf.writestr("other/file.xml", "<data/>")
        
        zip_buffer.seek(0)
        
        with patch('zipfile.ZipFile', side_effect=[zipfile.ZipFile(zip_buffer, 'r'), zipfile.ZipFile(BytesIO(), 'w')]), \
             patch('shutil.move'), \
             patch('tempfile.NamedTemporaryFile') as mock_temp:
            
            mock_temp.return_value.__enter__.return_value.name = "fixed.xlsx"
            
            # We can't easily test the result of the zip modification without a more complex setup,
            # but we can ensure the logic runs without error and returns the expected path.
            result_path = parser._fix_excel_styles("dummy.zip")
            assert result_path.endswith("_fixed.zip")

    def test_parse_with_corrupted_file(self):
        """测试解析损坏文件的处理。"""
        parser = XlsxParser()

        # 模拟样式相关的错误（包含"Fill"关键字）
        load_calls = 0
        def mock_load_side_effect(file_path, **kwargs):
            nonlocal load_calls
            load_calls += 1
            if load_calls <= 3:  # 前3次调用都失败（对应3种加载方式）
                raise Exception("Fill error in styles")
            elif load_calls == 4:  # 修复后的第一次调用成功
                # 修复后的调用成功
                mock_workbook = MagicMock()
                mock_worksheet = MagicMock()
                mock_worksheet.title = "FixedSheet"
                mock_worksheet.max_row = 1
                mock_worksheet.max_column = 1
                mock_cell = MagicMock()
                mock_cell.value = "Fixed"
                mock_worksheet.cell.return_value = mock_cell
                mock_worksheet.merged_cells.ranges = []
                mock_workbook.worksheets = [mock_worksheet]
                mock_workbook.sheetnames = ["FixedSheet"]
                mock_workbook.__getitem__.return_value = mock_worksheet
                return mock_workbook
            else:
                # data_only版本的调用
                return mock_workbook

        with patch('openpyxl.load_workbook', side_effect=mock_load_side_effect):
            with patch.object(parser, '_fix_excel_styles', return_value="fixed.xlsx"):
                sheets = parser.parse("corrupted.xlsx")
                assert len(sheets) == 1
                assert sheets[0].name == "FixedSheet"

    def test_parse_with_binary_io(self):
        """测试使用BinaryIO解析。"""
        parser = XlsxParser()

        # 创建一个模拟的BinaryIO对象
        mock_file = BytesIO(b"fake xlsx content")

        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_worksheet = MagicMock()
            mock_worksheet.title = "BinarySheet"
            mock_worksheet.max_row = 1
            mock_worksheet.max_column = 1
            mock_cell = MagicMock()
            mock_cell.value = "Binary"
            mock_worksheet.cell.return_value = mock_cell
            mock_worksheet.merged_cells.ranges = []
            mock_workbook.worksheets = [mock_worksheet]
            mock_workbook.sheetnames = ["BinarySheet"]
            mock_workbook.__getitem__.return_value = mock_worksheet
            mock_load.return_value = mock_workbook

            sheets = parser.parse(mock_file)
            assert len(sheets) == 1
            assert sheets[0].name == "BinarySheet"

    def test_parse_row_with_various_cell_types(self):
        """测试解析包含各种单元格类型的行。"""
        provider = XlsxRowProvider("test.xlsx")

        # 创建不同类型的模拟单元格
        text_cell = MagicMock(spec=OpenpyxlCell)
        text_cell.value = "Text"
        text_cell.data_type = 's'

        number_cell = MagicMock(spec=OpenpyxlCell)
        number_cell.value = 123
        number_cell.data_type = 'n'

        formula_cell = MagicMock(spec=OpenpyxlCell)
        formula_cell.value = "=SUM(A1:B1)"
        formula_cell.data_type = 'f'

        bool_cell = MagicMock(spec=OpenpyxlCell)
        bool_cell.value = True
        bool_cell.data_type = 'b'

        cells = (text_cell, number_cell, formula_cell, bool_cell)

        with patch('src.utils.style_parser.extract_cell_value', side_effect=lambda cell: cell.value):
            with patch('src.utils.style_parser.extract_style', return_value=None):
                row = provider._parse_row(cells)

                assert len(row.cells) == 4
                assert row.cells[0].value == "Text"
                assert row.cells[1].value == 123
                assert row.cells[2].value == "=SUM(A1:B1)"
                assert row.cells[3].value == True

    def test_extract_charts_with_different_types(self, tmp_path):
        """测试提取不同类型的图表。"""
        parser = XlsxParser()

        # 创建模拟的工作表和图表
        mock_worksheet = MagicMock()

        # 创建不同类型的图表
        bar_chart = MagicMock(spec=OpenpyxlBarChart)
        bar_chart.title = "Bar Chart"

        line_chart = MagicMock()
        line_chart.title = "Line Chart"

        pie_chart = MagicMock()
        pie_chart.title = "Pie Chart"

        mock_worksheet._charts = [bar_chart, line_chart, pie_chart]

        with patch('src.utils.chart_data_extractor.ChartDataExtractor') as mock_extractor_class:
            mock_extractor = MagicMock()
            mock_extractor_class.return_value = mock_extractor
            mock_extractor.extract_chart_data.return_value = {"type": "test", "data": []}

            charts = parser._extract_charts(mock_worksheet)

            assert len(charts) == 3
            assert all(isinstance(chart, Chart) for chart in charts)

    def test_extract_images_with_positioning(self):
        """测试提取带位置信息的图片。"""
        parser = XlsxParser()

        mock_worksheet = MagicMock()

        # 创建模拟图片
        mock_image = MagicMock(spec=OpenpyxlImage)
        mock_anchor = MagicMock()

        # 创建模拟的from_cell和to_cell
        mock_from_cell = MagicMock()
        mock_from_cell.col = 1
        mock_from_cell.row = 1
        mock_from_cell.colOff = 0
        mock_from_cell.rowOff = 0

        mock_to_cell = MagicMock()
        mock_to_cell.col = 3
        mock_to_cell.row = 3
        mock_to_cell.colOff = 0
        mock_to_cell.rowOff = 0

        mock_anchor._from = mock_from_cell
        mock_anchor._to = mock_to_cell

        mock_image.anchor = mock_anchor
        mock_image._data.return_value = b"image_data"

        mock_worksheet._images = [mock_image]

        with patch('src.utils.chart_data_extractor.ChartDataExtractor') as mock_extractor_class:
            mock_extractor = MagicMock()
            mock_extractor_class.return_value = mock_extractor
            mock_extractor.extract_chart_data.return_value = {"type": "image", "data": []}

            charts = parser._extract_images(mock_worksheet)

            assert len(charts) == 1
            chart = charts[0]
            assert chart.position.from_col == 1
            assert chart.position.from_row == 1
            assert chart.position.to_col == 3
            assert chart.position.to_row == 3

    def test_fix_excel_styles_no_styles_xml(self):
        """测试修复不包含styles.xml的Excel文件。"""
        parser = XlsxParser()

        # 模拟zipfile.ZipFile的行为
        with patch('zipfile.ZipFile') as mock_zipfile:
            mock_zip_instance = MagicMock()
            mock_zip_instance.namelist.return_value = ["other/file.xml"]  # 不包含styles.xml
            mock_zipfile.return_value.__enter__.return_value = mock_zip_instance

            # 如果没有styles.xml，方法仍然会创建修复后的文件
            result_path = parser._fix_excel_styles("dummy.xlsx")
            assert result_path.endswith("_fixed.xlsx")
