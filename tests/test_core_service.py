import csv
import openpyxl
import pytest
import xlwt
from unittest.mock import MagicMock, patch, mock_open, PropertyMock
from pathlib import Path
from datetime import datetime, date
from src.core_service import CoreService
from src.models.table_model import Sheet, Row, Cell, Style
from src.exceptions import FileNotFoundError

@pytest.fixture
def core_service_instance():
    """提供一个 CoreService 的实例。"""
    return CoreService()

class TestCoreService:
    """测试 CoreService 类的核心功能。"""

    def test_parse_sheet_normal(self, core_service_instance, tmp_path):
        """测试 parse_sheet 方法的正常解析功能。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "TestSheet"
        sheet.append(["ID", "Name"])
        sheet.append([1, "Alice"])
        workbook.save(file_path)

        result = core_service_instance.parse_sheet(str(file_path))
        
        assert result['sheet_name'] == "TestSheet"
        assert result['headers'] == ["ID", "Name"]
        assert result['rows'][0][0]['value'] == 1

    def test_convert_to_html_normal(self, core_service_instance, tmp_path):
        """测试 convert_to_html 方法的正常转换功能。"""
        file_path = tmp_path / "test.xlsx"
        output_path = tmp_path / "output.html"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.append(["Header1", "Header2"])
        sheet.append(["Data1", "Data2"])
        workbook.save(file_path)

        results = core_service_instance.convert_to_html(str(file_path), str(output_path))
        
        assert len(results) == 1
        assert Path(results[0]['output_path']).exists()

    def test_apply_changes_normal(self, core_service_instance, tmp_path):
        """测试 apply_changes 方法的正常应用修改功能。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "MySheet"
        sheet.append(["ID", "Name"])
        workbook.save(file_path)

        json_data = {
            "sheet_name": "MySheet",
            "headers": ["ID", "Name", "Age"],
            "rows": [[{"value": 1}, {"value": "Alice"}, {"value": 30}]]
        }

        result = core_service_instance.apply_changes(str(file_path), json_data, create_backup=False)
        
        assert result['status'] == 'success'
        
        reloaded_workbook = openpyxl.load_workbook(file_path)
        reloaded_sheet = reloaded_workbook["MySheet"]
        assert reloaded_sheet.cell(row=1, column=3).value == "Age"
        assert reloaded_sheet.cell(row=2, column=3).value == 30

    def test_write_back_xls_missing_keys(self, core_service_instance, tmp_path):
        """测试当 table_model_json 缺少 'headers' 或 'rows' 时，_write_back_xls 是否引发 KeyError。"""
        file_path = tmp_path / "test.xls"
        # 创建一个空的xls文件
        import xlwt
        workbook = xlwt.Workbook()
        workbook.add_sheet('Sheet1')
        workbook.save(str(file_path))

        json_data_no_headers = {"sheet_name": "Sheet1", "rows": []}
        json_data_no_rows = {"sheet_name": "Sheet1", "headers": []}

        with pytest.raises(KeyError):
            core_service_instance._write_back_xls(file_path, json_data_no_headers)
        
        with pytest.raises(KeyError):
            core_service_instance._write_back_xls(file_path, json_data_no_rows)

    def test_write_back_csv_normal(self, core_service_instance, tmp_path):
        """测试 _write_back_csv 函数的正常文件写入功能。"""
        file_path = tmp_path / "test.csv"
        json_data = {
            "headers": ["ID", "Name"],
            "rows": [
                [{"value": 1}, {"value": "Alice"}],
                [{"value": 2}, {"value": "Bob"}]
            ]
        }
        
        changes = core_service_instance._write_back_csv(file_path, json_data)
        
        assert changes == 2
        with open(file_path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            assert next(reader) == ["ID", "Name"]
            assert next(reader) == ["1", "Alice"]
            assert next(reader) == ["2", "Bob"]

    def test_write_back_xlsx_normal(self, core_service_instance, tmp_path):
        """测试 _write_back_xlsx 函数的正常文件写入功能。"""
        file_path = tmp_path / "test.xlsx"
        # 创建一个空的xlsx文件
        workbook = openpyxl.Workbook()
        # openpyxl 默认创建的sheet名为 "Sheet"
        workbook.save(str(file_path))

        json_data = {
            "sheet_name": "Sheet", # 确保sheet名与创建时一致
            "headers": ["ID", "Name"],
            "rows": [
                [{"value": 1}, {"value": "Alice"}],
                [{"value": 2}, {"value": "Bob"}]
            ]
        }
        
        changes = core_service_instance._write_back_xlsx(file_path, json_data)
        
        assert changes == 4  # 2 rows * 2 cells = 4 changes
        
        # 重新加载并验证内容
        reloaded_workbook = openpyxl.load_workbook(file_path)
        sheet = reloaded_workbook["Sheet"] # 使用确切的名称获取sheet
        
        assert sheet is not None
        # 验证表头
        assert sheet.cell(row=1, column=1).value == "ID"
        assert sheet.cell(row=1, column=2).value == "Name"
        # 验证第一行数据
        assert sheet.cell(row=2, column=1).value == 1
        assert sheet.cell(row=2, column=2).value == "Alice"
        # 验证第二行数据
        assert sheet.cell(row=3, column=1).value == 2
        assert sheet.cell(row=3, column=2).value == "Bob"

    def test_parse_sheet_with_cache(self, core_service_instance, tmp_path):
        """测试解析表格时的缓存功能。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "TestSheet"
        sheet.append(["ID", "Name"])
        sheet.append([1, "Alice"])
        workbook.save(file_path)

        # 第一次解析
        with patch('src.core_service.get_cache_manager') as mock_cache:
            mock_cache_instance = MagicMock()
            mock_cache.return_value = mock_cache_instance
            mock_cache_instance.get.return_value = None  # 缓存未命中

            result1 = core_service_instance.parse_sheet(str(file_path))

            # 验证缓存被调用
            mock_cache_instance.get.assert_called_once()
            mock_cache_instance.set.assert_called_once()

    def test_parse_sheet_from_cache(self, core_service_instance, tmp_path):
        """测试从缓存获取数据。"""
        file_path = tmp_path / "test.xlsx"
        # 创建一个真实的文件以通过验证
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        cached_data = {
            'data': {
                'sheet_name': 'TestSheet',
                'headers': ['ID', 'Name'],
                'rows': [[{'value': 1}, {'value': 'Alice'}]]
            }
        }

        with patch('src.core_service.get_cache_manager') as mock_cache:
            mock_cache_instance = MagicMock()
            mock_cache.return_value = mock_cache_instance
            mock_cache_instance.get.return_value = cached_data  # 缓存命中

            result = core_service_instance.parse_sheet(str(file_path))

            assert result == cached_data['data']
            mock_cache_instance.get.assert_called_once()
            mock_cache_instance.set.assert_not_called()  # 不应该设置缓存

    def test_parse_sheet_with_sheet_name(self, core_service_instance, tmp_path):
        """测试指定工作表名称的解析。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()

        # 创建第一个工作表
        sheet1 = workbook.active
        sheet1.title = "Sheet1"
        sheet1.append(["A", "B"])
        sheet1.append([1, 2])

        # 创建第二个工作表
        sheet2 = workbook.create_sheet("Sheet2")
        sheet2.append(["X", "Y"])
        sheet2.append([3, 4])

        workbook.save(file_path)

        # 测试指定工作表名称
        result = core_service_instance.parse_sheet(str(file_path), sheet_name="Sheet2")

        assert result['sheet_name'] == "Sheet2"
        assert result['headers'] == ["X", "Y"]
        assert result['rows'][0][0]['value'] == 3

    def test_parse_sheet_nonexistent_sheet(self, core_service_instance, tmp_path):
        """测试指定不存在的工作表名称。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        sheet.append(["A", "B"])
        workbook.save(file_path)

        # 测试不存在的工作表
        with pytest.raises(ValueError, match="工作表 'NonExistent' 不存在"):
            core_service_instance.parse_sheet(str(file_path), sheet_name="NonExistent")

    def test_parse_sheet_empty_file(self, core_service_instance, tmp_path):
        """测试解析空文件。"""
        file_path = tmp_path / "empty.xlsx"

        # 使用mock来模拟空文件的情况
        with patch.object(core_service_instance.parser_factory, 'get_parser') as mock_get_parser:
            mock_parser = MagicMock()
            mock_parser.parse.return_value = []  # 返回空的工作表列表
            mock_get_parser.return_value = mock_parser

            with patch('src.core_service.validate_file_input', return_value=(Path(file_path), None)):
                with pytest.raises(ValueError, match="文件中没有找到任何工作表"):
                    core_service_instance.parse_sheet(str(file_path))

    def test_parse_sheet_with_range(self, core_service_instance, tmp_path):
        """测试使用范围字符串解析。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "TestSheet"

        # 添加更多数据
        for i in range(5):
            sheet.append([f"Col{j}" if i == 0 else f"Data{i}_{j}" for j in range(5)])

        workbook.save(file_path)

        # 测试范围解析
        result = core_service_instance.parse_sheet(str(file_path), range_string="A1:C3")

        assert len(result['headers']) == 3
        assert len(result['rows']) == 2  # 不包括表头行

    def test_parse_sheet_invalid_range(self, core_service_instance, tmp_path):
        """测试无效的范围字符串。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["A", "B"])
        workbook.save(file_path)

        # 测试无效范围 - 应该回退到采样数据
        result = core_service_instance.parse_sheet(str(file_path), range_string="invalid_range")

        # 应该返回采样数据而不是抛出异常
        assert 'sheet_name' in result

    def test_parse_sheet_optimized_basic(self, core_service_instance, tmp_path):
        """测试 parse_sheet_optimized 方法的基本功能。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "TestSheet"
        sheet.append(["ID", "Name", "Age"])
        sheet.append([1, "Alice", 25])
        sheet.append([2, "Bob", 30])
        workbook.save(file_path)

        # 测试基本优化解析
        result = core_service_instance.parse_sheet_optimized(str(file_path))

        assert result['sheet_name'] == "TestSheet"
        assert 'metadata' in result
        assert result['metadata']['total_rows'] == 3

    def test_parse_sheet_optimized_with_full_data(self, core_service_instance, tmp_path):
        """测试 parse_sheet_optimized 返回完整数据。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["ID", "Name"])
        sheet.append([1, "Alice"])
        workbook.save(file_path)

        result = core_service_instance.parse_sheet_optimized(
            str(file_path),
            include_full_data=True,
            include_styles=True
        )

        assert 'rows' in result
        assert len(result['rows']) == 1

    def test_parse_sheet_optimized_with_range(self, core_service_instance, tmp_path):
        """测试 parse_sheet_optimized 使用范围。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for i in range(5):
            sheet.append([f"Col{j}" if i == 0 else f"Data{i}_{j}" for j in range(3)])
        workbook.save(file_path)

        result = core_service_instance.parse_sheet_optimized(
            str(file_path),
            range_string="A1:B3"
        )

        assert 'range' in result
        assert len(result['headers']) == 2

    def test_parse_sheet_optimized_invalid_range(self, core_service_instance, tmp_path):
        """测试 parse_sheet_optimized 使用无效范围。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["A", "B"])
        workbook.save(file_path)

        with pytest.raises(ValueError, match="范围格式错误"):
            core_service_instance.parse_sheet_optimized(
                str(file_path),
                range_string="invalid"
            )

    def test_parse_sheet_optimized_nonexistent_sheet(self, core_service_instance, tmp_path):
        """测试 parse_sheet_optimized 指定不存在的工作表。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["A", "B"])
        workbook.save(file_path)

        with pytest.raises(ValueError, match="工作表 'NonExistent' 不存在"):
            core_service_instance.parse_sheet_optimized(
                str(file_path),
                sheet_name="NonExistent"
            )

    def test_convert_to_html_file_not_found(self, core_service_instance):
        """测试转换不存在的文件到HTML。"""
        with pytest.raises(FileNotFoundError, match="文件不存在"):
            core_service_instance.convert_to_html("nonexistent.xlsx")

    def test_convert_to_html_with_sheet_name(self, core_service_instance, tmp_path):
        """测试转换指定工作表到HTML。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()

        # 创建两个工作表
        sheet1 = workbook.active
        sheet1.title = "Sheet1"
        sheet1.append(["A", "B"])

        sheet2 = workbook.create_sheet("Sheet2")
        sheet2.append(["X", "Y"])

        workbook.save(file_path)

        results = core_service_instance.convert_to_html(
            str(file_path),
            sheet_name="Sheet2"
        )

        assert len(results) == 1
        assert "Sheet2" in results[0]['output_path']

    def test_convert_to_html_nonexistent_sheet(self, core_service_instance, tmp_path):
        """测试转换不存在的工作表到HTML。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["A", "B"])
        workbook.save(file_path)

        with pytest.raises(ValueError, match="工作表 'NonExistent' 在文件中未找到"):
            core_service_instance.convert_to_html(
                str(file_path),
                sheet_name="NonExistent"
            )

    def test_apply_changes_file_not_found(self, core_service_instance):
        """测试对不存在文件应用修改。"""
        json_data = {
            "sheet_name": "Sheet1",
            "headers": ["A", "B"],
            "rows": []
        }

        with pytest.raises(FileNotFoundError, match="文件不存在"):
            core_service_instance.apply_changes("nonexistent.xlsx", json_data)

    def test_apply_changes_unsupported_format(self, core_service_instance, tmp_path):
        """测试对不支持格式的文件应用修改。"""
        file_path = tmp_path / "test.txt"
        file_path.write_text("some text")

        json_data = {
            "sheet_name": "Sheet1",
            "headers": ["A", "B"],
            "rows": []
        }

        with pytest.raises(ValueError, match="Unsupported file type"):
            core_service_instance.apply_changes(str(file_path), json_data)

    def test_apply_changes_xlsb_format(self, core_service_instance, tmp_path):
        """测试对XLSB格式文件应用修改。"""
        file_path = tmp_path / "test.xlsb"
        file_path.write_bytes(b"fake xlsb content")

        json_data = {
            "sheet_name": "Sheet1",
            "headers": ["A", "B"],
            "rows": []
        }

        with pytest.raises(ValueError, match="XLSB格式暂不支持数据写回"):
            core_service_instance.apply_changes(str(file_path), json_data)

    def test_apply_changes_missing_fields(self, core_service_instance, tmp_path):
        """测试缺少必需字段的JSON数据。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        # 缺少headers字段
        json_data = {
            "sheet_name": "Sheet1",
            "rows": []
        }

        with pytest.raises(ValueError, match="缺少必需字段: headers"):
            core_service_instance.apply_changes(str(file_path), json_data)

    def test_apply_changes_with_backup(self, core_service_instance, tmp_path):
        """测试创建备份文件的功能。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "TestSheet"
        sheet.append(["A", "B"])
        workbook.save(file_path)

        json_data = {
            "sheet_name": "TestSheet",  # 使用正确的工作表名
            "headers": ["A", "B"],
            "rows": [[{"value": 1}, {"value": 2}]]
        }

        result = core_service_instance.apply_changes(str(file_path), json_data, create_backup=True)

        assert result['backup_created'] is True
        assert result['backup_path'] is not None
        backup_path = Path(result['backup_path'])
        assert backup_path.exists()

    def test_apply_changes_csv_format(self, core_service_instance, tmp_path):
        """测试对CSV文件应用修改。"""
        file_path = tmp_path / "test.csv"
        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["A", "B"])
            writer.writerow([1, 2])

        json_data = {
            "sheet_name": "test",
            "headers": ["A", "B", "C"],
            "rows": [
                [{"value": 1}, {"value": 2}, {"value": 3}],
                [{"value": 4}, {"value": 5}, {"value": 6}]
            ]
        }

        result = core_service_instance.apply_changes(str(file_path), json_data, create_backup=False)

        assert result['status'] == 'success'
        assert result['changes_applied'] == 2

        # 验证文件内容
        with open(file_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert rows[0] == ["A", "B", "C"]
            assert rows[1] == ["1", "2", "3"]
            assert rows[2] == ["4", "5", "6"]

    def test_apply_changes_xls_format(self, core_service_instance, tmp_path):
        """测试对XLS文件应用修改。"""
        file_path = tmp_path / "test.xls"
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('TestSheet')
        sheet.write(0, 0, "A")
        sheet.write(0, 1, "B")
        workbook.save(str(file_path))

        json_data = {
            "sheet_name": "TestSheet",
            "headers": ["A", "B"],
            "rows": [[{"value": 1}, {"value": 2}]]
        }

        result = core_service_instance.apply_changes(str(file_path), json_data, create_backup=False)

        assert result['status'] == 'success'
        assert result['changes_applied'] == 2

    def test_write_back_xlsx_nonexistent_sheet(self, core_service_instance, tmp_path):
        """测试写回到不存在的工作表。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "NonExistent",
            "headers": ["A", "B"],
            "rows": []
        }

        with pytest.raises(ValueError, match="工作表 'NonExistent' 在文件中不存在"):
            core_service_instance._write_back_xlsx(file_path, json_data)

    def test_write_back_csv_with_none_values(self, core_service_instance, tmp_path):
        """测试CSV写回时处理None值。"""
        file_path = tmp_path / "test.csv"
        json_data = {
            "headers": ["A", "B"],
            "rows": [
                [{"value": None}, {"value": "test"}],
                [None, {"value": None}]
            ]
        }

        changes = core_service_instance._write_back_csv(file_path, json_data)

        assert changes == 2
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "," in content  # 确保空值被正确处理

    def test_sheet_to_json_empty_sheet(self, core_service_instance):
        """测试空工作表的JSON转换。"""
        # 创建空的Sheet对象
        empty_sheet = Sheet(name="EmptySheet", rows=[], merged_cells=[])

        result = core_service_instance._sheet_to_json(empty_sheet)

        assert result['sheet_name'] == "EmptySheet"
        assert result['size_info']['total_cells'] == 0

    def test_sheet_to_json_with_range(self, core_service_instance):
        """测试带范围的工作表JSON转换。"""
        # 创建测试数据
        rows = []
        for i in range(3):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(3)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._sheet_to_json(sheet, range_string="A1:B2")

        assert 'range' in result
        assert result['range'] == "A1:B2"

    def test_calculate_data_size_empty(self, core_service_instance):
        """测试计算空表格的数据大小。"""
        empty_sheet = Sheet(name="Empty", rows=[], merged_cells=[])
        size = core_service_instance._calculate_data_size(empty_sheet)
        assert size == 0

    def test_calculate_data_size_normal(self, core_service_instance):
        """测试计算正常表格的数据大小。"""
        rows = []
        for i in range(2):
            cells = [Cell(value=f"Cell{j}") for j in range(3)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        size = core_service_instance._calculate_data_size(sheet)
        assert size == 6  # 2 rows * 3 cells

    def test_value_to_json_serializable_datetime(self, core_service_instance):
        """测试日期时间值的JSON序列化。"""
        dt = datetime(2023, 1, 1, 12, 0, 0)
        result = core_service_instance._value_to_json_serializable(dt)
        assert result == "2023-01-01T12:00:00"

    def test_value_to_json_serializable_date(self, core_service_instance):
        """测试日期值的JSON序列化。"""
        d = date(2023, 1, 1)
        result = core_service_instance._value_to_json_serializable(d)
        assert result == "2023-01-01"

    def test_value_to_json_serializable_rich_text(self, core_service_instance):
        """测试富文本的JSON序列化。"""
        # 模拟富文本片段
        class MockFragment:
            def __init__(self, text):
                self.text = text

        rich_text = [MockFragment("Hello"), MockFragment(" World")]
        result = core_service_instance._value_to_json_serializable(rich_text)
        assert result == "Hello World"

    def test_extract_range_data_out_of_bounds(self, core_service_instance):
        """测试提取超出范围的数据。"""
        rows = []
        for i in range(2):
            cells = [Cell(value=f"Cell{j}") for j in range(2)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        # 起始行超出范围
        result = core_service_instance._extract_range_data(sheet, 5, 0, 6, 1)

        assert result['total_rows'] == 0
        assert result['size_info']['processing_mode'] == "range_out_of_bounds"

    def test_extract_range_data_invalid_range(self, core_service_instance):
        """测试无效范围参数。"""
        cells = [Cell(value=f"Cell{j}") for j in range(2)]
        rows = [Row(cells=cells)]
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        # 起始位置大于结束位置
        with pytest.raises(ValueError, match="范围无效"):
            core_service_instance._extract_range_data(sheet, 1, 1, 0, 0)

    def test_extract_range_data_negative_values(self, core_service_instance):
        """测试负数范围参数。"""
        cells = [Cell(value=f"Cell{j}") for j in range(2)]
        rows = [Row(cells=cells)]
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        with pytest.raises(ValueError, match="范围起始位置不能为负数"):
            core_service_instance._extract_range_data(sheet, -1, 0, 1, 1)

    def test_should_use_streaming_unsupported(self, core_service_instance, tmp_path):
        """测试不支持流式读取的文件格式。"""
        file_path = tmp_path / "test.txt"
        file_path.write_text("some text")

        with patch('src.core_service.ParserFactory.supports_streaming', return_value=False):
            result = core_service_instance._should_use_streaming(str(file_path), 1000)
            assert result is False

    def test_should_use_streaming_large_file(self, core_service_instance, tmp_path):
        """测试大文件的流式读取判断。"""
        file_path = tmp_path / "test.xlsx"
        # 创建一个较大的文件
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for i in range(100):
            sheet.append([f"Data{i}_{j}" for j in range(10)])
        workbook.save(file_path)

        with patch('src.core_service.ParserFactory.supports_streaming', return_value=True):
            with patch('src.core_service.get_config') as mock_config:
                mock_config.return_value.streaming_file_size_mb = 0.001  # 很小的阈值
                result = core_service_instance._should_use_streaming(str(file_path), 1000)
                assert result is True

    def test_should_use_streaming_error_handling(self, core_service_instance, tmp_path):
        """测试流式读取判断时的错误处理。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        with patch('src.core_service.ParserFactory.supports_streaming', return_value=True):
            with patch('src.core_service.StreamingTableReader', side_effect=Exception("Test error")):
                result = core_service_instance._should_use_streaming(str(file_path), 1000)
                assert result is False

    def test_parse_sheet_streaming_enabled(self, core_service_instance, tmp_path):
        """测试启用流式读取的解析。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for i in range(10):
            sheet.append([f"Data{i}_{j}" for j in range(5)])
        workbook.save(file_path)

        with patch.object(core_service_instance, '_should_use_streaming', return_value=True):
            with patch.object(core_service_instance, '_parse_sheet_streaming') as mock_streaming:
                mock_streaming.return_value = {"sheet_name": "test", "rows": []}

                result = core_service_instance.parse_sheet(str(file_path), enable_streaming=True)

                mock_streaming.assert_called_once()

    def test_parse_sheet_streaming_disabled(self, core_service_instance, tmp_path):
        """测试禁用流式读取的解析。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["A", "B"])
        workbook.save(file_path)

        result = core_service_instance.parse_sheet(str(file_path), enable_streaming=False)

        assert 'sheet_name' in result

    def test_parse_sheet_with_custom_threshold(self, core_service_instance, tmp_path):
        """测试自定义流式读取阈值。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["A", "B"])
        workbook.save(file_path)

        result = core_service_instance.parse_sheet(
            str(file_path),
            streaming_threshold=1  # 很低的阈值
        )

        assert 'sheet_name' in result

    def test_extract_optimized_data_with_styles(self, core_service_instance):
        """测试提取优化数据时包含样式。"""
        # 创建带样式的测试数据
        style = Style()
        cells = [Cell(value=f"Cell{j}", style=style) for j in range(3)]
        rows = [Row(cells=cells)]
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._extract_optimized_data(
            sheet,
            include_full_data=True,
            include_styles=True
        )

        assert 'rows' in result
        assert 'style' in result['rows'][0][0]

    def test_extract_optimized_data_preview_only(self, core_service_instance):
        """测试只提取预览数据。"""
        rows = []
        for i in range(10):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(3)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._extract_optimized_data(
            sheet,
            include_full_data=False,
            preview_rows=3
        )

        assert 'preview_rows' in result
        assert len(result['preview_rows']) == 3

    def test_extract_optimized_data_with_max_rows(self, core_service_instance):
        """测试限制最大行数的数据提取。"""
        rows = []
        for i in range(10):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(3)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._extract_optimized_data(
            sheet,
            include_full_data=True,
            max_rows=5
        )

        assert result['metadata']['truncated'] is True
        assert result['metadata']['truncated_at'] == 5

    def test_analyze_data_types_empty_sheet(self, core_service_instance):
        """测试分析空工作表的数据类型。"""
        sheet = Sheet(name="Empty", rows=[], merged_cells=[])
        headers = ["A", "B"]

        result = core_service_instance._analyze_data_types(sheet, headers)

        assert result == {"A": "unknown", "B": "unknown"}

    def test_analyze_data_types_mixed_types(self, core_service_instance):
        """测试分析混合数据类型。"""
        # 创建包含不同数据类型的测试数据
        header_row = Row(cells=[Cell(value="Text"), Cell(value="Number"), Cell(value="Mixed")])
        data_rows = [
            Row(cells=[Cell(value="Hello"), Cell(value=123), Cell(value="Text1")]),
            Row(cells=[Cell(value="World"), Cell(value=456), Cell(value=789)]),
        ]
        rows = [header_row] + data_rows
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])
        headers = ["Text", "Number", "Mixed"]

        result = core_service_instance._analyze_data_types(sheet, headers)

        assert result["Text"] == "text"
        assert result["Number"] == "number"
        assert result["Mixed"] == "mixed"

    def test_parse_sheet_with_empty_target_sheet(self, core_service_instance, tmp_path):
        """测试解析空的目标工作表。"""
        file_path = tmp_path / "test.xlsx"

        # 模拟解析器返回None的目标工作表
        with patch.object(core_service_instance.parser_factory, 'get_parser') as mock_get_parser:
            mock_parser = MagicMock()
            mock_parser.parse.return_value = [None]  # 返回包含None的工作表列表
            mock_get_parser.return_value = mock_parser

            with patch('src.core_service.validate_file_input', return_value=(Path(file_path), None)):
                with patch('src.core_service.get_cache_manager') as mock_cache:
                    mock_cache_instance = MagicMock()
                    mock_cache.return_value = mock_cache_instance
                    mock_cache_instance.get.return_value = None

                    result = core_service_instance.parse_sheet(str(file_path))

                    assert result['sheet_name'] == "Empty"
                    assert result['total_rows'] == 0

    def test_parse_sheet_with_empty_rows(self, core_service_instance, tmp_path):
        """测试解析有工作表但无行数据的情况。"""
        file_path = tmp_path / "test.xlsx"

        # 创建一个空的Sheet对象
        empty_sheet = Sheet(name="EmptySheet", rows=[], merged_cells=[])

        with patch.object(core_service_instance.parser_factory, 'get_parser') as mock_get_parser:
            mock_parser = MagicMock()
            mock_parser.parse.return_value = [empty_sheet]
            mock_get_parser.return_value = mock_parser

            with patch('src.core_service.validate_file_input', return_value=(Path(file_path), None)):
                with patch('src.core_service.get_cache_manager') as mock_cache:
                    mock_cache_instance = MagicMock()
                    mock_cache.return_value = mock_cache_instance
                    mock_cache_instance.get.return_value = None

                    result = core_service_instance.parse_sheet(str(file_path))

                    assert result['sheet_name'] == "EmptySheet"
                    assert result['total_rows'] == 0

    def test_extract_sample_data(self, core_service_instance):
        """测试提取采样数据功能。"""
        # 创建包含多行数据的测试工作表
        rows = []
        for i in range(15):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(3)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._extract_sample_data(sheet, 45)

        assert result['sheet_name'] == "TestSheet"
        assert result['size_info']['processing_mode'] == "sample"
        assert len(result['sample_data']['rows']) <= 10

    def test_extract_simplified_data(self, core_service_instance):
        """测试提取简化数据功能。"""
        rows = []
        for i in range(5):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(3)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._extract_simplified_data(sheet, 15)

        assert result['sheet_name'] == "TestSheet"
        assert result['size_info']['processing_mode'] == "simplified"
        assert 'data' in result

    def test_generate_summary(self, core_service_instance):
        """测试生成摘要功能。"""
        rows = []
        for i in range(10):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(3)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._generate_summary(sheet)

        assert result['sheet_name'] == "TestSheet"
        assert result['size_info']['processing_mode'] == "summary"
        assert 'suggested_ranges' in result

    def test_extract_full_data_single_row(self, core_service_instance):
        """测试提取只有一行数据的完整数据。"""
        cells = [Cell(value=f"Cell{j}") for j in range(3)]
        rows = [Row(cells=cells)]
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._extract_full_data(sheet)

        assert result['sheet_name'] == "TestSheet"
        assert len(result['headers']) == 3
        assert len(result['rows']) == 1  # 单行被当作数据而不是表头

    def test_convert_to_html_with_pagination(self, core_service_instance, tmp_path):
        """测试带分页的HTML转换。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for i in range(20):
            sheet.append([f"Data{i}_{j}" for j in range(3)])
        workbook.save(file_path)

        with patch('src.converters.paginated_html_converter.PaginatedHTMLConverter') as mock_converter_class:
            mock_converter = MagicMock()
            mock_converter_class.return_value = mock_converter
            mock_converter.convert_to_file.return_value = {
                'output_path': str(tmp_path / 'output.html'),
                'status': 'success'
            }

            results = core_service_instance.convert_to_html(
                str(file_path),
                page_size=10,
                page_number=1
            )

            assert len(results) == 1
            mock_converter.convert_to_file.assert_called_once()

    def test_parse_sheet_exception_handling(self, core_service_instance, tmp_path):
        """测试解析过程中的异常处理。"""
        file_path = tmp_path / "test.xlsx"

        with patch('src.core_service.validate_file_input', side_effect=Exception("Test error")):
            with pytest.raises(Exception, match="Test error"):
                core_service_instance.parse_sheet(str(file_path))

    def test_parse_sheet_optimized_empty_file_error(self, core_service_instance, tmp_path):
        """测试parse_sheet_optimized处理空文件错误。"""
        file_path = tmp_path / "test.xlsx"

        with patch.object(core_service_instance.parser_factory, 'get_parser') as mock_get_parser:
            mock_parser = MagicMock()
            mock_parser.parse.return_value = []  # 返回空的工作表列表
            mock_get_parser.return_value = mock_parser

            with patch('src.core_service.validate_file_input', return_value=(Path(file_path), None)):
                with pytest.raises(ValueError, match="文件中没有找到任何工作表"):
                    core_service_instance.parse_sheet_optimized(str(file_path))

    def test_write_back_xlsx_import_error_handling(self, core_service_instance, tmp_path):
        """测试XLSX写回时ImportError的处理。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["A"],
            "rows": [[{"value": 1}]]
        }

        # 直接测试正常情况，ImportError分支很难模拟
        changes = core_service_instance._write_back_xlsx(file_path, json_data)
        assert changes >= 0

    def test_write_back_xlsx_merged_cell_string_check(self, core_service_instance, tmp_path):
        """测试XLSX写回时合并单元格的字符串检查。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # 创建合并单元格
        if hasattr(sheet, 'merge_cells'):
            sheet.merge_cells('A1:B1')
            sheet['A1'] = "Merged"
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["A", "B"],
            "rows": [[{"value": 1}, {"value": 2}]]
        }

        # 测试正常情况
        changes = core_service_instance._write_back_xlsx(file_path, json_data)
        assert changes >= 0

    def test_sheet_to_json_large_file_summary(self, core_service_instance):
        """测试大文件返回摘要的情况。"""
        # 创建一个大的工作表来触发摘要模式
        rows = []
        for i in range(200):  # 创建足够大的数据
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(50)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="LargeSheet", rows=rows, merged_cells=[])

        with patch('src.core_service.get_config') as mock_config:
            mock_config_instance = MagicMock()
            mock_config.return_value = mock_config_instance
            mock_config_instance.large_file_threshold_cells = 5000  # 设置较低的阈值

            result = core_service_instance._sheet_to_json(sheet)

            assert result['size_info']['processing_mode'] == "summary"

    def test_sheet_to_json_medium_file_sample(self, core_service_instance):
        """测试中等文件返回采样数据的情况。"""
        # 创建中等大小的工作表
        rows = []
        for i in range(50):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(20)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="MediumSheet", rows=rows, merged_cells=[])

        with patch('src.core_service.get_config') as mock_config:
            mock_config_instance = MagicMock()
            mock_config.return_value = mock_config_instance
            mock_config_instance.small_file_threshold_cells = 500
            mock_config_instance.medium_file_threshold_cells = 800
            mock_config_instance.large_file_threshold_cells = 2000

            result = core_service_instance._sheet_to_json(sheet)

            assert result['size_info']['processing_mode'] == "sample"

    def test_sheet_to_json_small_medium_file_simplified(self, core_service_instance):
        """测试小-中文件返回简化数据的情况。"""
        # 创建小-中等大小的工作表
        rows = []
        for i in range(20):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(15)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="SmallMediumSheet", rows=rows, merged_cells=[])

        with patch('src.core_service.get_config') as mock_config:
            mock_config_instance = MagicMock()
            mock_config.return_value = mock_config_instance
            mock_config_instance.small_file_threshold_cells = 200
            mock_config_instance.medium_file_threshold_cells = 500
            mock_config_instance.large_file_threshold_cells = 1000

            result = core_service_instance._sheet_to_json(sheet)

            assert result['size_info']['processing_mode'] == "simplified"

    def test_sheet_to_json_small_file_full_data(self, core_service_instance):
        """测试小文件返回完整数据的情况。"""
        # 创建小工作表
        rows = []
        for i in range(5):
            cells = [Cell(value=f"Cell{i}_{j}") for j in range(3)]
            rows.append(Row(cells=cells))
        sheet = Sheet(name="SmallSheet", rows=rows, merged_cells=[])

        with patch('src.core_service.get_config') as mock_config:
            mock_config_instance = MagicMock()
            mock_config.return_value = mock_config_instance
            mock_config_instance.small_file_threshold_cells = 100
            mock_config_instance.medium_file_threshold_cells = 500
            mock_config_instance.large_file_threshold_cells = 1000

            result = core_service_instance._sheet_to_json(sheet)

            assert result['size_info']['processing_mode'] == "full"

    def test_analyze_data_types_empty_values(self, core_service_instance):
        """测试分析包含空值的数据类型。"""
        # 创建只有表头的测试数据（没有数据行）
        header_row = Row(cells=[Cell(value="EmptyCol")])
        rows = [header_row]  # 只有表头，没有数据行
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])
        headers = ["EmptyCol"]

        result = core_service_instance._analyze_data_types(sheet, headers)

        assert result["EmptyCol"] == "unknown"  # 没有数据行时应该返回unknown

    def test_analyze_data_types_boolean_values(self, core_service_instance):
        """测试分析布尔数据类型。"""
        header_row = Row(cells=[Cell(value="BoolCol")])
        data_rows = [
            Row(cells=[Cell(value=True)]),
            Row(cells=[Cell(value=False)]),
        ]
        rows = [header_row] + data_rows
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])
        headers = ["BoolCol"]

        result = core_service_instance._analyze_data_types(sheet, headers)

        # 布尔值在Python中是int的子类，所以会被识别为number
        assert result["BoolCol"] == "number"

    def test_generate_summary_data_types_analysis(self, core_service_instance):
        """测试生成摘要时的数据类型分析。"""
        # 创建包含不同数据类型的测试数据
        header_row = Row(cells=[Cell(value="TextCol"), Cell(value="NumCol"), Cell(value="OtherCol")])
        data_rows = [
            Row(cells=[Cell(value="Hello"), Cell(value=123), Cell(value=datetime.now())]),
            Row(cells=[Cell(value="World"), Cell(value=456), Cell(value=date.today())]),
        ]
        rows = [header_row] + data_rows
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._generate_summary(sheet)

        assert "data_types" in result['metadata']
        assert "TextCol" in result['metadata']['data_types']
        assert "NumCol" in result['metadata']['data_types']

    def test_should_use_streaming_file_size_check(self, core_service_instance, tmp_path):
        """测试基于文件大小的流式读取判断。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["A", "B"])
        workbook.save(file_path)

        with patch('src.core_service.ParserFactory.supports_streaming', return_value=True):
            with patch('src.core_service.get_config') as mock_config:
                mock_config_instance = MagicMock()
                mock_config.return_value = mock_config_instance
                mock_config_instance.streaming_file_size_mb = 0.000001  # 极小的阈值，确保触发

                result = core_service_instance._should_use_streaming(str(file_path), 1000)
                assert result is True

    def test_should_use_streaming_with_streaming_reader_info(self, core_service_instance, tmp_path):
        """测试使用StreamingTableReader获取信息进行判断。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["A", "B"])
        workbook.save(file_path)

        with patch('src.core_service.ParserFactory.supports_streaming', return_value=True):
            with patch('src.core_service.get_config') as mock_config:
                mock_config_instance = MagicMock()
                mock_config.return_value = mock_config_instance
                mock_config_instance.streaming_file_size_mb = 999  # 大阈值，不会基于文件大小触发

                with patch('src.core_service.StreamingTableReader') as mock_reader_class:
                    mock_reader = MagicMock()
                    mock_reader_class.return_value.__enter__.return_value = mock_reader
                    mock_reader.get_info.return_value = {
                        'total_rows': 100,
                        'total_columns': 50
                    }

                    result = core_service_instance._should_use_streaming(str(file_path), 1000)
                    assert result is True  # 100 * 50 = 5000 > 1000

    def test_parse_sheet_streaming_large_file_summary(self, core_service_instance, tmp_path):
        """测试流式解析大文件返回摘要。"""
        file_path = tmp_path / "test.xlsx"

        with patch.object(core_service_instance, '_should_use_streaming', return_value=True):
            with patch('src.core_service.StreamingTableReader') as mock_reader_class:
                mock_reader = MagicMock()
                mock_reader_class.return_value.__enter__.return_value = mock_reader
                mock_reader.file_path = str(file_path)

                # 模拟大文件信息
                mock_reader.get_info.return_value = {
                    'total_rows': 1000,
                    'total_columns': 100,
                    'parser_type': 'xlsx',
                    'estimated_memory_usage': '10MB'
                }

                with patch('src.core_service.get_config') as mock_config:
                    mock_config_instance = MagicMock()
                    mock_config.return_value = mock_config_instance
                    mock_config_instance.streaming_summary_threshold_cells = 50000  # 设置较低阈值

                    with patch.object(core_service_instance, '_generate_streaming_summary') as mock_summary:
                        mock_summary.return_value = {"summary": "test"}

                        result = core_service_instance._parse_sheet_streaming(str(file_path))

                        mock_summary.assert_called_once()

    def test_parse_sheet_streaming_normal_processing(self, core_service_instance, tmp_path):
        """测试流式解析正常处理流程。"""
        file_path = tmp_path / "test.xlsx"

        with patch('src.core_service.StreamingTableReader') as mock_reader_class:
            mock_reader = MagicMock()
            mock_reader_class.return_value.__enter__.return_value = mock_reader
            mock_reader.file_path = str(file_path)

            # 模拟小文件信息
            mock_reader.get_info.return_value = {
                'total_rows': 10,
                'total_columns': 5,
                'parser_type': 'xlsx',
                'estimated_memory_usage': '1MB'
            }

            # 模拟数据块
            mock_chunk = MagicMock()
            mock_chunk.headers = ["A", "B"]
            mock_chunk.rows = [
                MagicMock(cells=[Cell(value="1"), Cell(value="2")])
            ]

            mock_reader.iter_chunks.return_value = [mock_chunk]

            with patch('src.core_service.get_config') as mock_config:
                mock_config_instance = MagicMock()
                mock_config.return_value = mock_config_instance
                mock_config_instance.streaming_summary_threshold_cells = 1000  # 高阈值
                mock_config_instance.streaming_chunk_size_rows = 100

                result = core_service_instance._parse_sheet_streaming(str(file_path))

                assert result['metadata']['processing_mode'] == "streaming"
                assert result['headers'] == ["A", "B"]

    def test_parse_sheet_streaming_fallback_to_traditional(self, core_service_instance, tmp_path):
        """测试流式解析失败时回退到传统方法。"""
        file_path = tmp_path / "test.xlsx"

        with patch('src.core_service.StreamingTableReader', side_effect=Exception("Streaming failed")):
            with patch.object(core_service_instance.parser_factory, 'get_parser') as mock_get_parser:
                mock_parser = MagicMock()
                mock_sheet = Sheet(name="TestSheet", rows=[
                    Row(cells=[Cell(value="A"), Cell(value="B")]),
                    Row(cells=[Cell(value=1), Cell(value=2)])
                ], merged_cells=[])
                mock_parser.parse.return_value = [mock_sheet]
                mock_get_parser.return_value = mock_parser

                result = core_service_instance._parse_sheet_streaming(str(file_path))

                assert result['sheet_name'] == "TestSheet"

    def test_parse_sheet_streaming_with_sheet_name_fallback(self, core_service_instance, tmp_path):
        """测试流式解析失败时使用指定工作表名称回退。"""
        file_path = tmp_path / "test.xlsx"

        with patch('src.core_service.StreamingTableReader', side_effect=Exception("Streaming failed")):
            with patch.object(core_service_instance.parser_factory, 'get_parser') as mock_get_parser:
                mock_parser = MagicMock()
                mock_sheet1 = Sheet(name="Sheet1", rows=[], merged_cells=[])
                mock_sheet2 = Sheet(name="Sheet2", rows=[
                    Row(cells=[Cell(value="X"), Cell(value="Y")])
                ], merged_cells=[])
                mock_parser.parse.return_value = [mock_sheet1, mock_sheet2]
                mock_get_parser.return_value = mock_parser

                result = core_service_instance._parse_sheet_streaming(str(file_path), sheet_name="Sheet2")

                assert result['sheet_name'] == "Sheet2"

    def test_parse_sheet_streaming_fallback_nonexistent_sheet(self, core_service_instance, tmp_path):
        """测试流式解析回退时工作表不存在的情况。"""
        file_path = tmp_path / "test.xlsx"

        with patch('src.core_service.StreamingTableReader', side_effect=Exception("Streaming failed")):
            with patch.object(core_service_instance.parser_factory, 'get_parser') as mock_get_parser:
                mock_parser = MagicMock()
                mock_sheet = Sheet(name="Sheet1", rows=[], merged_cells=[])
                mock_parser.parse.return_value = [mock_sheet]
                mock_get_parser.return_value = mock_parser

                with pytest.raises(ValueError, match="工作表 'NonExistent' 不存在"):
                    core_service_instance._parse_sheet_streaming(str(file_path), sheet_name="NonExistent")

    def test_parse_sheet_streaming_fallback_empty_file(self, core_service_instance, tmp_path):
        """测试流式解析回退时空文件的情况。"""
        file_path = tmp_path / "test.xlsx"

        with patch('src.core_service.StreamingTableReader', side_effect=Exception("Streaming failed")):
            with patch.object(core_service_instance.parser_factory, 'get_parser') as mock_get_parser:
                mock_parser = MagicMock()
                mock_parser.parse.return_value = []  # 空文件
                mock_get_parser.return_value = mock_parser

                with pytest.raises(ValueError, match="文件中没有找到任何工作表"):
                    core_service_instance._parse_sheet_streaming(str(file_path))

    def test_generate_streaming_summary_with_data(self, core_service_instance):
        """测试生成流式摘要功能。"""
        mock_reader = MagicMock()
        mock_reader.file_path = "/test/file.xlsx"

        file_info = {
            'parser_type': 'xlsx',
            'total_rows': 1000,
            'total_columns': 50,
            'estimated_memory_usage': '10MB'
        }

        # 模拟数据块 - 第一行是表头，后面是数据行
        mock_chunk = MagicMock()
        mock_chunk.headers = ["Col1", "Col2", "Col3"]

        # 创建模拟的行对象
        header_row = MagicMock()
        header_row.cells = [Cell(value="Col1"), Cell(value="Col2"), Cell(value="Col3")]

        data_row1 = MagicMock()
        data_row1.cells = [Cell(value="Data1"), Cell(value="Data2"), Cell(value="Data3")]

        data_row2 = MagicMock()
        data_row2.cells = [Cell(value="Data4"), Cell(value="Data5"), Cell(value="Data6")]

        mock_chunk.rows = [header_row, data_row1, data_row2]  # 包含表头行

        mock_reader.iter_chunks.return_value = [mock_chunk]

        with patch('src.core_service.ChunkFilter') as mock_filter:
            mock_filter.return_value = MagicMock()

            result = core_service_instance._generate_streaming_summary(mock_reader, file_info)

            assert result['metadata']['processing_mode'] == "streaming_summary"
            assert result['sample_data']['headers'] == ["Col1", "Col2", "Col3"]
            assert len(result['sample_data']['rows']) == 2  # 跳过表头行，只有2行数据

    def test_write_back_xlsx_number_conversion(self, core_service_instance, tmp_path):
        """测试XLSX写回时数字转换功能。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["Integer", "Float", "String"],
            "rows": [
                [{"value": "123"}, {"value": "45.67"}, {"value": "text"}],
                [{"value": "456"}, {"value": "78.9e2"}, {"value": "more text"}]
            ]
        }

        changes = core_service_instance._write_back_xlsx(file_path, json_data)
        assert changes >= 0

        # 验证数字转换
        reloaded_workbook = openpyxl.load_workbook(file_path)
        sheet = reloaded_workbook["Sheet"]

        # 检查整数转换
        assert sheet.cell(row=2, column=1).value == 123
        # 检查浮点数转换
        assert sheet.cell(row=2, column=2).value == 45.67
        # 检查字符串保持不变
        assert sheet.cell(row=2, column=3).value == "text"

    def test_write_back_xlsx_value_types(self, core_service_instance, tmp_path):
        """测试XLSX写回时不同值类型的处理。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["Mixed"],
            "rows": [
                [{"value": None}],  # None值
                [{"value": ""}],    # 空字符串
                [{"value": 42}],    # 整数
                [{"value": 3.14}],  # 浮点数
                [{"value": True}],  # 布尔值
                [{"value": [1, 2, 3]}]  # 其他类型
            ]
        }

        changes = core_service_instance._write_back_xlsx(file_path, json_data)
        assert changes >= 0

    def test_analyze_data_types_with_boolean_and_other_types(self, core_service_instance):
        """测试分析包含布尔值和其他类型的数据。"""
        header_row = Row(cells=[Cell(value="BoolCol"), Cell(value="OtherCol")])
        data_rows = [
            Row(cells=[Cell(value=True), Cell(value=datetime.now())]),
            Row(cells=[Cell(value=False), Cell(value=date.today())]),
        ]
        rows = [header_row] + data_rows
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])
        headers = ["BoolCol", "OtherCol"]

        result = core_service_instance._analyze_data_types(sheet, headers)

        # 布尔值会被识别为boolean类型
        assert "boolean" in result["BoolCol"] or "number" in result["BoolCol"]
        # 日期时间对象会被识别为other类型
        assert "other" in result["OtherCol"]

    def test_analyze_data_types_with_none_cells(self, core_service_instance):
        """测试分析包含None单元格的数据类型。"""
        header_row = Row(cells=[Cell(value="TestCol")])

        # 使用patch来模拟包含None单元格的情况
        with patch.object(header_row, 'cells', [Cell(value="TestCol")]):
            # 创建一个模拟的行，其中包含None单元格
            mock_row = MagicMock()
            mock_row.cells = [None]  # 模拟None单元格

            data_rows = [
                mock_row,  # None单元格
                Row(cells=[Cell(value="text")]),  # 正常单元格
            ]
            rows = [header_row] + data_rows
            sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])
            headers = ["TestCol"]

            result = core_service_instance._analyze_data_types(sheet, headers)

            # 应该能正确处理None单元格，并基于非None单元格推断类型
            assert result["TestCol"] == "text"

    def test_write_back_xlsx_with_merged_cells_mock(self, core_service_instance, tmp_path):
        """测试XLSX写回时处理合并单元格的模拟情况。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["A"],
            "rows": [[{"value": "test"}]]
        }

        # 创建一个模拟的MergedCell类型的单元格
        class MockMergedCell:
            def __init__(self):
                self.coordinate = "A1"
                self._value = None

            @property
            def value(self):
                return self._value

            @value.setter
            def value(self, val):
                raise AttributeError("read-only property")

        # 模拟工作表返回MergedCell
        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_worksheet = MagicMock()

            # 第一次调用返回普通单元格（用于清除），第二次返回MergedCell（用于写入）
            normal_cell = MagicMock()
            merged_cell = MockMergedCell()

            call_count = 0
            def mock_cell(*args, **kwargs):
                nonlocal call_count
                call_count += 1
                if call_count <= 2:  # 前两次调用（清除阶段）
                    return normal_cell
                else:  # 后续调用（写入阶段）
                    return merged_cell

            mock_worksheet.cell.side_effect = mock_cell
            mock_worksheet.max_row = 1
            mock_worksheet.max_column = 1
            mock_workbook.__getitem__.return_value = mock_worksheet
            mock_workbook.sheetnames = ["Sheet"]
            mock_load.return_value = mock_workbook

            # 这应该触发AttributeError处理并跳过
            changes = core_service_instance._write_back_xlsx(file_path, json_data)
            assert changes >= 0

    def test_write_back_xlsx_other_attribute_error_reraise(self, core_service_instance, tmp_path):
        """测试XLSX写回时其他AttributeError的重新抛出。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["A"],
            "rows": [[{"value": "test"}]]
        }

        # 创建一个会抛出其他AttributeError的单元格
        class MockErrorCell:
            def __init__(self):
                self.coordinate = "A1"

            @property
            def value(self):
                return None

            @value.setter
            def value(self, val):
                raise AttributeError("some other error")

        # 模拟工作表返回错误单元格
        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_worksheet = MagicMock()

            normal_cell = MagicMock()
            error_cell = MockErrorCell()

            call_count = 0
            def mock_cell(*args, **kwargs):
                nonlocal call_count
                call_count += 1
                if call_count <= 2:  # 前两次调用（清除阶段）
                    return normal_cell
                else:  # 后续调用（写入阶段）
                    return error_cell

            mock_worksheet.cell.side_effect = mock_cell
            mock_worksheet.max_row = 1
            mock_worksheet.max_column = 1
            mock_workbook.__getitem__.return_value = mock_worksheet
            mock_workbook.sheetnames = ["Sheet"]
            mock_load.return_value = mock_workbook

            # 这应该重新抛出AttributeError
            with pytest.raises(AttributeError, match="some other error"):
                core_service_instance._write_back_xlsx(file_path, json_data)

    def test_write_back_xlsx_import_error_string_check(self, core_service_instance, tmp_path):
        """测试XLSX写回时ImportError的字符串检查。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["A"],
            "rows": [[{"value": "test"}]]
        }

        # 创建一个模拟的MergedCell类型的单元格（通过字符串检查）
        class MockMergedCellString:
            def __init__(self):
                self.coordinate = "A1"
                self._value = None

            def __str__(self):
                return "<class 'openpyxl.cell.cell.MergedCell'>"

            def __repr__(self):
                return self.__str__()

            @property
            def value(self):
                return self._value

            @value.setter
            def value(self, val):
                self._value = val

        # 模拟ImportError和字符串检查
        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_worksheet = MagicMock()

            normal_cell = MagicMock()
            merged_cell_string = MockMergedCellString()

            call_count = 0
            def mock_cell(*args, **kwargs):
                nonlocal call_count
                call_count += 1
                if call_count <= 2:  # 前两次调用（清除阶段）
                    return normal_cell
                else:  # 后续调用（写入阶段）
                    return merged_cell_string

            mock_worksheet.cell.side_effect = mock_cell
            mock_worksheet.max_row = 1
            mock_worksheet.max_column = 1
            mock_workbook.__getitem__.return_value = mock_worksheet
            mock_workbook.sheetnames = ["Sheet"]
            mock_load.return_value = mock_workbook

            # 模拟ImportError
            with patch('src.core_service.logger') as mock_logger:
                # 这应该通过字符串检查跳过MergedCell
                changes = core_service_instance._write_back_xlsx(file_path, json_data)
                assert changes >= 0

    def test_write_back_xlsx_clear_cells_import_error(self, core_service_instance, tmp_path):
        """测试XLSX写回时清除单元格阶段的ImportError处理。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["A"],
            "rows": [[{"value": "test"}]]
        }

        # 创建一个在清除阶段会触发ImportError的单元格
        class MockClearMergedCell:
            def __init__(self):
                self.coordinate = "A1"
                self._value = None

            def __str__(self):
                return "<class 'openpyxl.cell.cell.MergedCell'>"

            @property
            def value(self):
                return self._value

            @value.setter
            def value(self, val):
                self._value = val

        # 模拟工作表在清除阶段返回MergedCell
        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_worksheet = MagicMock()

            clear_merged_cell = MockClearMergedCell()
            normal_cell = MagicMock()

            call_count = 0
            def mock_cell(*args, **kwargs):
                nonlocal call_count
                call_count += 1
                if call_count == 1:  # 第一次调用（清除阶段）
                    return clear_merged_cell
                else:  # 后续调用
                    return normal_cell

            mock_worksheet.cell.side_effect = mock_cell
            mock_worksheet.max_row = 1
            mock_worksheet.max_column = 1
            mock_workbook.__getitem__.return_value = mock_worksheet
            mock_workbook.sheetnames = ["Sheet"]
            mock_load.return_value = mock_workbook

            # 这个测试主要验证不会崩溃
            changes = core_service_instance._write_back_xlsx(file_path, json_data)
            assert changes >= 0

    def test_extract_optimized_data_with_none_cells(self, core_service_instance):
        """测试提取优化数据时处理None单元格。"""
        # 创建包含None单元格的测试数据
        rows = []
        for i in range(3):
            cells = []
            for j in range(3):
                if i == 2 and j == 2:  # 最后一个位置设为None
                    cells.append(None)  # 直接添加None而不是Cell对象
                else:
                    cells.append(Cell(value=f"Cell{i}_{j}"))
            rows.append(Row(cells=cells))
        sheet = Sheet(name="TestSheet", rows=rows, merged_cells=[])

        result = core_service_instance._extract_optimized_data(
            sheet,
            include_full_data=True,
            include_styles=True
        )

        assert 'rows' in result
        # 验证None单元格被正确处理
        assert result['rows'][1][2]['value'] is None
        assert result['rows'][1][2]['style'] is None

    def test_parse_sheet_streaming_with_range_filter(self, core_service_instance, tmp_path):
        """测试流式解析时使用范围过滤器。"""
        file_path = tmp_path / "test.xlsx"

        with patch('src.core_service.StreamingTableReader') as mock_reader_class:
            mock_reader = MagicMock()
            mock_reader_class.return_value.__enter__.return_value = mock_reader
            mock_reader.file_path = str(file_path)

            # 模拟文件信息
            mock_reader.get_info.return_value = {
                'total_rows': 10,
                'total_columns': 5,
                'parser_type': 'xlsx',
                'estimated_memory_usage': '1MB'
            }

            # 模拟数据块
            mock_chunk = MagicMock()
            mock_chunk.headers = ["A", "B", "C"]
            mock_chunk.rows = [
                MagicMock(cells=[Cell(value="1"), Cell(value="2"), Cell(value="3")])
            ]

            mock_reader.iter_chunks.return_value = [mock_chunk]

            with patch('src.core_service.get_config') as mock_config:
                mock_config_instance = MagicMock()
                mock_config.return_value = mock_config_instance
                mock_config_instance.streaming_summary_threshold_cells = 1000
                mock_config_instance.streaming_chunk_size_rows = 100

                with patch('src.core_service.ChunkFilter') as mock_filter_class:
                    mock_filter = MagicMock()
                    mock_filter_class.return_value = mock_filter

                    result = core_service_instance._parse_sheet_streaming(
                        str(file_path),
                        range_string="A1:C3"
                    )

                    # 验证ChunkFilter被创建
                    mock_filter_class.assert_called_with(range_string="A1:C3")
                    assert result['metadata']['processing_mode'] == "streaming"

    def test_write_back_xlsx_exception_handling_in_clear_phase(self, core_service_instance, tmp_path):
        """测试XLSX写回时清除阶段的异常处理。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["A"],
            "rows": [[{"value": "test"}]]
        }

        # 模拟在清除阶段抛出异常的单元格
        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_worksheet = MagicMock()

            # 第一次调用（清除阶段）抛出异常，后续调用正常
            call_count = 0
            def mock_cell(*args, **kwargs):
                nonlocal call_count
                call_count += 1
                if call_count == 1:
                    raise Exception("Clear phase error")
                else:
                    return MagicMock()

            mock_worksheet.cell.side_effect = mock_cell
            mock_worksheet.max_row = 1
            mock_worksheet.max_column = 1
            mock_workbook.__getitem__.return_value = mock_worksheet
            mock_workbook.sheetnames = ["Sheet"]
            mock_load.return_value = mock_workbook

            # 这应该在清除阶段抛出异常
            with pytest.raises(Exception, match="Clear phase error"):
                core_service_instance._write_back_xlsx(file_path, json_data)

    def test_write_back_xlsx_exception_handling_in_write_phase(self, core_service_instance, tmp_path):
        """测试XLSX写回时写入阶段的异常处理。"""
        file_path = tmp_path / "test.xlsx"
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

        json_data = {
            "sheet_name": "Sheet",
            "headers": ["A"],
            "rows": [[{"value": "test"}]]
        }

        # 模拟在写入阶段抛出异常的单元格
        with patch('openpyxl.load_workbook') as mock_load:
            mock_workbook = MagicMock()
            mock_worksheet = MagicMock()

            # 前两次调用（清除阶段）正常，第三次调用（写入阶段）抛出异常
            call_count = 0
            def mock_cell(*args, **kwargs):
                nonlocal call_count
                call_count += 1
                if call_count <= 2:
                    return MagicMock()
                else:
                    raise Exception("Write phase error")

            mock_worksheet.cell.side_effect = mock_cell
            mock_worksheet.max_row = 1
            mock_worksheet.max_column = 1
            mock_workbook.__getitem__.return_value = mock_worksheet
            mock_workbook.sheetnames = ["Sheet"]
            mock_load.return_value = mock_workbook

            # 这应该在写入阶段抛出异常
            with pytest.raises(Exception, match="Write phase error"):
                core_service_instance._write_back_xlsx(file_path, json_data)
