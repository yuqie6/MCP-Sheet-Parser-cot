import pytest
from pathlib import Path
from src.parsers.csv_parser import CsvParser, CsvRowProvider
from src.models.table_model import Sheet, LazySheet

@pytest.fixture
def create_csv_file(tmp_path: Path):
    """工厂fixture，用于创建不同编码和内容的CSV文件。"""
    files_created = []
    def _create_file(filename: str, content: str, encoding: str = 'utf-8'):
        file_path = tmp_path / filename
        file_path.write_text(content, encoding=encoding)
        files_created.append(file_path)
        return file_path
    yield _create_file
    # 清理创建的文件
    for file_path in files_created:
        if file_path.exists():
            file_path.unlink()

class TestCsvParser:
    """测试 CsvParser 类的功能。"""

    def test_parse_utf8_file(self, create_csv_file):
        """测试解析一个标准的UTF-8编码的CSV文件。"""
        content = "header1,header2\nvalue1,value2"
        file_path = create_csv_file("test_utf8.csv", content, "utf-8")
        
        parser = CsvParser()
        sheets = parser.parse(str(file_path))
        
        assert len(sheets) == 1
        sheet = sheets[0]
        assert isinstance(sheet, Sheet)
        assert sheet.name == "test_utf8"
        assert len(sheet.rows) == 2
        assert sheet.rows[0].cells[0].value == "header1"
        assert sheet.rows[1].cells[1].value == "value2"

    def test_parse_gbk_file(self, create_csv_file):
        """测试当UTF-8解码失败时，能否成功回退到GBK编码。"""
        content = "标题1,标题2\n值1,值2"
        file_path = create_csv_file("test_gbk.csv", content, "gbk")

        parser = CsvParser()
        sheets = parser.parse(str(file_path))

        assert len(sheets) == 1
        sheet = sheets[0]
        assert sheet.name == "test_gbk"
        assert len(sheet.rows) == 2
        assert sheet.rows[0].cells[0].value == "标题1"
        assert sheet.rows[1].cells[1].value == "值2"

    def test_file_not_found(self):
        """测试当文件不存在时是否会抛出FileNotFoundError。"""
        parser = CsvParser()
        with pytest.raises(FileNotFoundError):
            parser.parse("non_existent_file.csv")

    def test_supports_streaming(self):
        """测试解析器是否正确报告其支持流式处理。"""
        parser = CsvParser()
        assert parser.supports_streaming() is True

    def test_create_lazy_sheet(self, create_csv_file):
        """测试创建LazySheet对象的功能，并验证其内容。"""
        content = "a,b\nc,d"
        file_path = create_csv_file("lazy.csv", content)
        
        parser = CsvParser()
        lazy_sheet = parser.create_lazy_sheet(str(file_path))
        
        assert isinstance(lazy_sheet, LazySheet)
        assert lazy_sheet.name == "lazy"
        
        # 不直接访问 provider，而是通过公共API验证行为
        assert lazy_sheet.get_total_rows() == 2
        row = lazy_sheet.get_row(1)
        assert row.cells[0].value == "c"
        assert row.cells[1].value == "d"

class TestCsvRowProvider:
    """测试 CsvRowProvider 类的功能。"""

    def test_get_total_rows(self, create_csv_file):
        """测试获取总行数的功能。"""
        content = "row1\nrow2\nrow3"
        file_path = create_csv_file("total_rows.csv", content)
        provider = CsvRowProvider(str(file_path))
        
        assert provider.get_total_rows() == 3
        # 测试缓存
        assert provider.get_total_rows() == 3

    def test_get_row(self, create_csv_file):
        """测试按索引获取指定行的功能。"""
        content = "a,b\nc,d\ne,f"
        file_path = create_csv_file("get_row.csv", content)
        provider = CsvRowProvider(str(file_path))
        
        row = provider.get_row(1)
        assert row.cells[0].value == "c"
        assert row.cells[1].value == "d"

    def test_get_row_out_of_bounds(self, create_csv_file):
        """测试当行索引超出范围时是否抛出IndexError。"""
        content = "a,b"
        file_path = create_csv_file("out_of_bounds.csv", content)
        provider = CsvRowProvider(str(file_path))
        
        with pytest.raises(IndexError):
            provider.get_row(5)

    def test_iter_rows_full(self, create_csv_file):
        """测试完整迭代所有行。"""
        content = "1,2\n3,4"
        file_path = create_csv_file("iter_full.csv", content)
        provider = CsvRowProvider(str(file_path))
        
        rows = list(provider.iter_rows())
        assert len(rows) == 2
        assert rows[0].cells[0].value == "1"
        assert rows[1].cells[1].value == "4"

    def test_iter_rows_with_start_row(self, create_csv_file):
        """测试从指定行开始迭代。"""
        content = "a\nb\nc\nd"
        file_path = create_csv_file("iter_start.csv", content)
        provider = CsvRowProvider(str(file_path))
        
        rows = list(provider.iter_rows(start_row=2))
        assert len(rows) == 2
        assert rows[0].cells[0].value == "c"

    def test_iter_rows_with_max_rows(self, create_csv_file):
        """测试迭代指定最大行数。"""
        content = "a\nb\nc\nd"
        file_path = create_csv_file("iter_max.csv", content)
        provider = CsvRowProvider(str(file_path))
        
        rows = list(provider.iter_rows(max_rows=2))
        assert len(rows) == 2
        assert rows[0].cells[0].value == "a"
        assert rows[1].cells[0].value == "b"

    def test_iter_rows_with_start_and_max(self, create_csv_file):
        """测试同时使用start_row和max_rows参数。"""
        content = "a\nb\nc\nd\ne"
        file_path = create_csv_file("iter_combo.csv", content)
        provider = CsvRowProvider(str(file_path))
        
        rows = list(provider.iter_rows(start_row=1, max_rows=3))
        assert len(rows) == 3
        assert rows[0].cells[0].value == "b"
        assert rows[2].cells[0].value == "d"

    def test_iter_rows_empty_file(self, create_csv_file):
        """测试迭代一个空文件。"""
        file_path = create_csv_file("empty.csv", "")
        provider = CsvRowProvider(str(file_path))
        
        rows = list(provider.iter_rows())
        assert len(rows) == 0