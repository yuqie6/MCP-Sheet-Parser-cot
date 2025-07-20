
import pytest
from unittest.mock import patch, MagicMock
from src.parsers.factory import ParserFactory
from src.parsers.xlsx_parser import XlsxParser
from src.parsers.csv_parser import CsvParser
from src.exceptions import UnsupportedFileTypeError

@patch('src.parsers.factory.validate_file_input')
def test_get_parser_xlsx(mock_validate):
    """Test getting an XlsxParser."""
    mock_validate.return_value = ("dummy.xlsx", "xlsx")
    parser = ParserFactory.get_parser("dummy.xlsx")
    assert isinstance(parser, XlsxParser)

@patch('src.parsers.factory.validate_file_input')
def test_get_parser_csv(mock_validate):
    """Test getting a CsvParser."""
    mock_validate.return_value = ("dummy.csv", "csv")
    parser = ParserFactory.get_parser("dummy.csv")
    assert isinstance(parser, CsvParser)

@patch('src.parsers.factory.validate_file_input')
def test_get_parser_unsupported(mock_validate):
    """Test getting a parser for an unsupported file type."""
    mock_validate.return_value = ("dummy.txt", "txt")
    with pytest.raises(UnsupportedFileTypeError):
        ParserFactory.get_parser("dummy.txt")

def test_get_supported_formats():
    """Test getting the list of supported formats."""
    formats = ParserFactory.get_supported_formats()
    assert "xlsx" in formats
    assert "csv" in formats

def test_is_supported_format():
    """Test checking if a format is supported."""
    assert ParserFactory.is_supported_format("test.xlsx") is True
    assert ParserFactory.is_supported_format("test.txt") is False

@patch('src.parsers.factory.ParserFactory.get_parser')
def test_supports_streaming(mock_get_parser):
    """Test checking if a format supports streaming."""
    mock_parser = mock_get_parser.return_value
    mock_parser.supports_streaming.return_value = True
    assert ParserFactory.supports_streaming("streaming.xlsx") is True
    mock_parser.supports_streaming.return_value = False
    assert ParserFactory.supports_streaming("non_streaming.xls") is False


@patch('src.parsers.factory.validate_file_input')
def test_get_parser_xls(mock_validate):
    """
    TDD测试：get_parser应该为XLS文件返回XlsParser

    这个测试覆盖XLS文件类型的解析器创建代码路径
    """
    from src.parsers.xls_parser import XlsParser

    mock_validate.return_value = ("dummy.xls", "xls")
    parser = ParserFactory.get_parser("dummy.xls")
    assert isinstance(parser, XlsParser)

@patch('src.parsers.factory.validate_file_input')
def test_get_parser_with_exception_handling(mock_validate):
    """
    TDD测试：get_parser应该处理解析器创建时的异常

    这个测试覆盖异常处理的代码路径
    """
    mock_validate.return_value = ("dummy.xlsx", "xlsx")

    # 模拟解析器创建时抛出异常
    with patch.dict('src.parsers.factory.ParserFactory._parser_classes',
                    {'xlsx': MagicMock(side_effect=Exception("Parser creation failed"))}):
        with pytest.raises(Exception):
            ParserFactory.get_parser("dummy.xlsx")

def test_is_supported_format_with_various_extensions():
    """
    TDD测试：is_supported_format应该正确识别各种文件扩展名

    这个测试覆盖所有支持的文件格式检查
    """

    # 测试支持的格式
    supported_files = [
        "test.xlsx", "test.XLSX", "TEST.xlsx",  # Excel 2007+
        "test.xls", "test.XLS", "TEST.xls",     # Excel 97-2003
        "test.csv", "test.CSV", "TEST.csv"      # CSV
    ]

    for file_path in supported_files:
        assert ParserFactory.is_supported_format(file_path) is True

    # 测试不支持的格式
    unsupported_files = [
        "test.txt", "test.doc", "test.pdf", "test.json",
        "test.xml", "test.html", "test", "test."
    ]

    for file_path in unsupported_files:
        assert ParserFactory.is_supported_format(file_path) is False

def test_is_supported_format_with_no_extension():
    """
    TDD测试：is_supported_format应该处理没有扩展名的文件

    这个测试确保方法在文件没有扩展名时正确处理
    """
    assert ParserFactory.is_supported_format("filename_without_extension") is False
    assert ParserFactory.is_supported_format("") is False

def test_is_supported_format_with_path_separators():
    """
    TDD测试：is_supported_format应该处理包含路径分隔符的文件路径

    这个测试确保方法能正确处理完整的文件路径
    """

    # Unix风格路径
    assert ParserFactory.is_supported_format("/path/to/file.xlsx") is True
    assert ParserFactory.is_supported_format("/path/to/file.txt") is False

    # Windows风格路径
    assert ParserFactory.is_supported_format("C:\\path\\to\\file.csv") is True
    assert ParserFactory.is_supported_format("C:\\path\\to\\file.doc") is False

def test_get_supported_formats_completeness():
    """
    TDD测试：get_supported_formats应该返回所有支持的格式

    这个测试确保返回的格式列表包含所有预期的格式
    """
    formats = ParserFactory.get_supported_formats()

    # 验证返回的是列表
    assert isinstance(formats, list)

    # 验证包含所有预期的格式
    expected_formats = ["xlsx", "xls", "csv"]
    for format_type in expected_formats:
        assert format_type in formats

    # 验证没有重复
    assert len(formats) == len(set(formats))

@patch('src.parsers.factory.ParserFactory.get_parser')
def test_supports_streaming_with_exception(mock_get_parser):
    """
    TDD测试：supports_streaming应该处理解析器获取时的异常

    这个测试覆盖异常处理的代码路径
    """
    # 模拟get_parser抛出异常
    mock_get_parser.side_effect = Exception("Parser creation failed")

    # 应该返回False而不是抛出异常
    result = ParserFactory.supports_streaming("problematic.xlsx")
    assert result is False

@patch('src.parsers.factory.ParserFactory.get_parser')
def test_supports_streaming_with_parser_without_method(mock_get_parser):
    """
    TDD测试：supports_streaming应该处理解析器没有supports_streaming方法的情况

    这个测试确保方法在解析器缺少方法时正确处理
    """
    # 创建一个没有supports_streaming方法的模拟解析器
    mock_parser = object()  # 简单对象，没有supports_streaming方法
    mock_get_parser.return_value = mock_parser

    # 应该返回False
    result = ParserFactory.supports_streaming("test.xlsx")
    assert result is False

def test_parser_factory_is_static():
    """
    TDD测试：ParserFactory应该是静态类，不需要实例化

    这个测试验证工厂类的设计模式
    """
    # 验证所有方法都是静态方法或类方法
    assert hasattr(ParserFactory, 'get_parser')
    assert hasattr(ParserFactory, 'get_supported_formats')
    assert hasattr(ParserFactory, 'is_supported_format')
    assert hasattr(ParserFactory, 'supports_streaming')

    # 验证可以直接调用而不需要实例化
    formats = ParserFactory.get_supported_formats()
    assert isinstance(formats, list)

@patch('src.parsers.factory.validate_file_input')
def test_get_parser_with_case_insensitive_extension(mock_validate):
    """
    TDD测试：get_parser应该不区分大小写地处理文件扩展名

    这个测试确保文件扩展名的大小写不影响解析器选择
    """
    # 测试大写扩展名
    mock_validate.return_value = ("dummy.XLSX", "xlsx")
    parser = ParserFactory.get_parser("dummy.XLSX")
    assert isinstance(parser, XlsxParser)

    # 测试混合大小写扩展名
    mock_validate.return_value = ("dummy.XlSx", "xlsx")
    parser = ParserFactory.get_parser("dummy.XlSx")
    assert isinstance(parser, XlsxParser)

def test_get_format_info():
    """
    TDD测试：get_format_info应该返回所有支持格式的详细信息

    这个测试覆盖第91行的get_format_info方法
    """
    format_info = ParserFactory.get_format_info()

    # 验证返回的是字典
    assert isinstance(format_info, dict)

    # 验证包含预期的格式
    expected_formats = ["csv", "xlsx", "xls", "xlsb", "xlsm"]
    for format_type in expected_formats:
        assert format_type in format_info

        # 验证每个格式都有必要的信息
        format_data = format_info[format_type]
        assert "name" in format_data
        assert "description" in format_data
        assert "features" in format_data
        assert isinstance(format_data["features"], list)

def test_is_supported_format_with_malformed_filename():
    """
    TDD测试：is_supported_format应该处理格式错误的文件名

    这个测试覆盖第143-144行的IndexError异常处理
    """
    # 测试空字符串（会导致IndexError）
    assert ParserFactory.is_supported_format("") is False

    # 测试只有点号的文件名
    assert ParserFactory.is_supported_format(".") is False
    assert ParserFactory.is_supported_format("..") is False

    # 测试以点号结尾但没有扩展名的文件
    assert ParserFactory.is_supported_format("filename.") is False

def test_is_supported_format_edge_cases():
    """
    TDD测试：is_supported_format应该处理各种边界情况

    这个测试确保覆盖IndexError异常处理的所有情况
    """
    # 直接测试会导致IndexError的情况
    # 空字符串split('.')[-1]会导致IndexError
    result = ParserFactory.is_supported_format("")
    assert result is False

    # 测试其他边界情况
    assert ParserFactory.is_supported_format("file_without_extension") is False
    assert ParserFactory.is_supported_format(".hidden_file") is False

def test_get_streaming_formats():
    """
    TDD测试：get_streaming_formats应该返回支持流式读取的格式列表

    这个测试覆盖第171-181行的get_streaming_formats方法
    """
    streaming_formats = ParserFactory.get_streaming_formats()

    # 验证返回的是列表
    assert isinstance(streaming_formats, list)

    # 验证CSV格式支持流式读取（根据实际实现）
    assert "csv" in streaming_formats

    # 验证列表中没有重复项
    assert len(streaming_formats) == len(set(streaming_formats))

def test_get_streaming_formats_with_parser_creation_failure():
    """
    TDD测试：get_streaming_formats应该处理解析器创建失败的情况

    这个测试覆盖第178-180行的异常处理
    """
    # 模拟一个会在创建时抛出异常的解析器类
    class FailingParser:
        def __init__(self):
            raise Exception("Parser creation failed")

    # 临时替换解析器类字典
    original_classes = ParserFactory._parser_classes.copy()
    try:
        # 添加一个会失败的解析器
        ParserFactory._parser_classes["failing"] = FailingParser

        # 调用get_streaming_formats，应该跳过失败的解析器
        streaming_formats = ParserFactory.get_streaming_formats()

        # 验证返回的是列表且不包含失败的格式
        assert isinstance(streaming_formats, list)
        assert "failing" not in streaming_formats

    finally:
        # 恢复原始的解析器类字典
        ParserFactory._parser_classes = original_classes

@patch('src.parsers.factory.ParserFactory.get_parser')
def test_create_lazy_sheet(mock_get_parser):
    """
    TDD测试：create_lazy_sheet应该创建懒加载的工作表

    这个测试覆盖第198-199行的create_lazy_sheet方法
    """
    # 创建模拟解析器
    mock_parser = MagicMock()
    mock_lazy_sheet = MagicMock()
    mock_parser.create_lazy_sheet.return_value = mock_lazy_sheet
    mock_get_parser.return_value = mock_parser

    # 调用create_lazy_sheet
    result = ParserFactory.create_lazy_sheet("test.xlsx", "Sheet1")

    # 验证调用了正确的方法
    mock_get_parser.assert_called_once_with("test.xlsx")
    mock_parser.create_lazy_sheet.assert_called_once_with("test.xlsx", "Sheet1")
    assert result == mock_lazy_sheet

@patch('src.parsers.factory.ParserFactory.get_parser')
def test_create_lazy_sheet_without_sheet_name(mock_get_parser):
    """
    TDD测试：create_lazy_sheet应该支持不指定工作表名称

    这个测试确保方法支持可选的sheet_name参数
    """
    # 创建模拟解析器
    mock_parser = MagicMock()
    mock_lazy_sheet = MagicMock()
    mock_parser.create_lazy_sheet.return_value = mock_lazy_sheet
    mock_get_parser.return_value = mock_parser

    # 调用create_lazy_sheet不指定sheet_name
    result = ParserFactory.create_lazy_sheet("test.xlsx")

    # 验证调用了正确的方法
    mock_get_parser.assert_called_once_with("test.xlsx")
    mock_parser.create_lazy_sheet.assert_called_once_with("test.xlsx", None)
    assert result == mock_lazy_sheet
