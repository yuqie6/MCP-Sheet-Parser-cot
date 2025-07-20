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


    def test_parse_with_encoding_detection_failure(self, create_csv_file):
        """
        TDD测试：parse应该处理编码检测失败的情况

        这个测试覆盖第28-29行的编码检测失败代码路径
        """
        # 创建一个包含特殊字符的文件，可能导致编码检测困难
        content = "header1,header2\nvalue1,value2"
        file_path = create_csv_file("encoding_test.csv", content, "latin-1")

        parser = CsvParser()

        # 应该能够解析，即使编码检测可能不完美
        sheets = parser.parse(str(file_path))
        assert len(sheets) == 1
        assert isinstance(sheets[0], Sheet)

    def test_parse_with_csv_error_handling(self, create_csv_file):
        """
        TDD测试：parse应该处理CSV解析错误

        这个测试覆盖第40-41行的CSV错误处理代码路径
        """
        # 创建一个格式错误的CSV文件
        content = 'header1,header2\n"unclosed quote,value2\nvalue3,value4'
        file_path = create_csv_file("malformed.csv", content)

        parser = CsvParser()

        # 应该能够处理错误并继续解析
        sheets = parser.parse(str(file_path))
        assert len(sheets) == 1

    def test_parse_with_io_error(self, tmp_path):
        """
        TDD测试：parse应该处理文件IO错误

        这个测试确保方法在文件不存在时正确处理
        """
        parser = CsvParser()
        non_existent_file = str(tmp_path / "non_existent.csv")

        # 应该抛出适当的异常
        with pytest.raises((FileNotFoundError, IOError)):
            parser.parse(non_existent_file)

    def test_supports_streaming(self):
        """
        TDD测试：CsvParser应该支持流式处理

        这个测试验证流式处理支持
        """
        parser = CsvParser()
        assert parser.supports_streaming() is True

    def test_create_lazy_sheet(self, create_csv_file):
        """
        TDD测试：create_lazy_sheet应该创建LazySheet对象

        这个测试覆盖第84行的LazySheet创建代码路径
        """
        content = "header1,header2\nvalue1,value2"
        file_path = create_csv_file("lazy_test.csv", content)

        parser = CsvParser()
        lazy_sheet = parser.create_lazy_sheet(str(file_path))

        assert lazy_sheet is not None
        assert isinstance(lazy_sheet, LazySheet)
        assert lazy_sheet.name == "lazy_test"

    def test_create_lazy_sheet_with_sheet_name(self, create_csv_file):
        """
        TDD测试：create_lazy_sheet应该处理sheet_name参数

        这个测试确保sheet_name参数被正确处理
        """
        content = "header1,header2\nvalue1,value2"
        file_path = create_csv_file("named_sheet.csv", content)

        parser = CsvParser()
        lazy_sheet = parser.create_lazy_sheet(str(file_path), "CustomName")

        assert lazy_sheet is not None
        assert lazy_sheet.name == "CustomName"

class TestCsvRowProviderAdditional:
    """额外的CsvRowProvider测试，提升覆盖率。"""

    def test_get_total_rows_with_empty_file(self, create_csv_file):
        """
        TDD测试：get_total_rows应该处理空文件

        这个测试确保空文件的行数计算正确
        """
        file_path = create_csv_file("empty_rows.csv", "")
        provider = CsvRowProvider(str(file_path))

        total_rows = provider.get_total_rows()
        assert total_rows == 0

    def test_get_total_rows_with_single_line(self, create_csv_file):
        """
        TDD测试：get_total_rows应该正确计算单行文件

        这个测试确保单行文件的行数计算正确
        """
        content = "header1,header2"
        file_path = create_csv_file("single_line.csv", content)
        provider = CsvRowProvider(str(file_path))

        total_rows = provider.get_total_rows()
        assert total_rows == 1

    def test_get_row_beyond_file_end(self, create_csv_file):
        """
        TDD测试：get_row应该处理超出文件末尾的行索引

        这个测试确保方法在索引超出范围时正确处理
        """
        content = "header1,header2\nvalue1,value2"
        file_path = create_csv_file("short_file.csv", content)
        provider = CsvRowProvider(str(file_path))

        # 尝试获取不存在的行，应该抛出IndexError
        with pytest.raises(IndexError, match="行索引 10 超出范围"):
            provider.get_row(10)

    def test_iter_rows_with_max_rows_exceeding_file(self, create_csv_file):
        """
        TDD测试：iter_rows应该处理max_rows超过文件行数的情况

        这个测试确保方法在请求的行数超过文件实际行数时正确处理
        """
        content = "a,b\nc,d"
        file_path = create_csv_file("short_iter.csv", content)
        provider = CsvRowProvider(str(file_path))

        # 请求比文件实际行数更多的行
        rows = list(provider.iter_rows(max_rows=100))
        assert len(rows) == 2  # 只应该返回实际存在的行数

    def test_iter_rows_with_start_row_at_end(self, create_csv_file):
        """
        TDD测试：iter_rows应该处理start_row在文件末尾的情况

        这个测试确保方法在起始行在文件末尾时返回空结果
        """
        content = "a,b\nc,d"
        file_path = create_csv_file("end_start.csv", content)
        provider = CsvRowProvider(str(file_path))

        # 从文件末尾开始迭代
        rows = list(provider.iter_rows(start_row=10))
        assert len(rows) == 0


class TestCsvRowProviderEncodingDetection:
    """测试CsvRowProvider的编码检测功能。"""

    def test_detect_encoding_with_unicode_decode_error(self, create_csv_file):
        """
        TDD测试：_detect_encoding应该处理UnicodeDecodeError并回退到GBK

        这个测试覆盖第28-29行的异常处理代码
        """

        # 创建一个包含GBK特有字符的文件，这些字符在UTF-8下会导致解码错误
        content = "测试,数据\n中文,内容"
        file_path = create_csv_file("test_gbk_encoding.csv", content, "gbk")

        # 创建CsvRowProvider实例，这会触发编码检测
        provider = CsvRowProvider(str(file_path))

        # 验证编码被正确检测为gbk
        assert provider._encoding == "gbk"

        # 验证能够正确读取内容
        rows = list(provider.iter_rows())
        assert len(rows) == 2
        assert rows[0].cells[0].value == "测试"
        assert rows[1].cells[1].value == "内容"

class TestCsvParserStyleExtraction:
    """测试CsvParser的样式提取功能。"""

    def test_extract_style_returns_none(self, create_csv_file):
        """
        TDD测试：_extract_style应该始终返回None（CSV不支持样式）

        这个测试覆盖第84行的返回None代码
        """

        content = "header1,header2\nvalue1,value2"
        file_path = create_csv_file("test_style.csv", content, "utf-8")

        parser = CsvParser()

        # 测试_extract_style方法直接调用
        result = parser._extract_style("any_cell_value")
        assert result is None

        # 测试通过解析验证样式确实为None
        sheets = parser.parse(str(file_path))
        sheet = sheets[0]

        # 验证所有单元格的样式都是None
        for row in sheet.rows:
            for cell in row.cells:
                assert cell.style is None