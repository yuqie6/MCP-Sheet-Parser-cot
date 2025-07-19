import pytest
from src.exceptions import (
    SheetParserError, FileError, FileNotFoundError, FileAccessError,
    UnsupportedFileTypeError, CorruptedFileError, ParseError, SheetNotFoundError,
    InvalidRangeError, StyleExtractionError, ConversionError, HTMLConversionError,
    ValidationError, ResourceError, MemoryLimitExceededError, FileSizeLimitExceededError,
    TimeoutError, CacheError, ConfigurationError, ERROR_CODE_MAP, get_error_by_code
)

# === TDD测试：提升Exceptions覆盖率到100% ===

class TestSheetParserError:
    """测试基础异常类SheetParserError。"""

    def test_basic_initialization(self):
        """
        TDD测试：SheetParserError应该正确初始化基本属性
        
        这个测试覆盖第17-29行的初始化代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        error = SheetParserError("Test error message")
        
        assert error.message == "Test error message"
        assert error.error_code == "SheetParserError"
        assert error.details == {}
        assert str(error) == "Test error message"

    def test_initialization_with_all_parameters(self):
        """
        TDD测试：SheetParserError应该处理所有参数
        
        这个测试确保所有参数都被正确设置
        """
        # 🔴 红阶段：编写测试描述期望的行为
        details = {"key": "value", "number": 123}
        error = SheetParserError("Custom message", "CUSTOM_CODE", details)
        
        assert error.message == "Custom message"
        assert error.error_code == "CUSTOM_CODE"
        assert error.details == details

    def test_to_dict_method(self):
        """
        TDD测试：to_dict应该返回正确的字典格式
        
        这个测试覆盖第31-38行的to_dict方法代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        details = {"test": "data"}
        error = SheetParserError("Test message", "TEST_CODE", details)
        
        result = error.to_dict()
        
        expected = {
            "error_type": "SheetParserError",
            "error_code": "TEST_CODE",
            "message": "Test message",
            "details": details
        }
        
        assert result == expected

class TestFileErrors:
    """测试文件相关的异常类。"""

    def test_file_not_found_error(self):
        """
        TDD测试：FileNotFoundError应该正确格式化错误信息
        
        这个测试覆盖第46-54行的FileNotFoundError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        file_path = "/path/to/missing/file.xlsx"
        error = FileNotFoundError(file_path)
        
        assert "文件不存在" in error.message
        assert file_path in error.message
        assert error.error_code == "FILE_NOT_FOUND"
        assert error.details["file_path"] == file_path

    def test_file_access_error_default_operation(self):
        """
        TDD测试：FileAccessError应该使用默认操作类型
        
        这个测试覆盖第57-65行的FileAccessError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        file_path = "/path/to/protected/file.xlsx"
        error = FileAccessError(file_path)
        
        assert "无法read文件" in error.message
        assert error.error_code == "FILE_ACCESS_DENIED"
        assert error.details["operation"] == "read"

    def test_file_access_error_custom_operation(self):
        """
        TDD测试：FileAccessError应该处理自定义操作类型
        
        这个测试确保自定义操作类型被正确处理
        """
        # 🔴 红阶段：编写测试描述期望的行为
        file_path = "/path/to/file.xlsx"
        error = FileAccessError(file_path, "write")
        
        assert "无法write文件" in error.message
        assert error.details["operation"] == "write"

    def test_unsupported_file_type_error(self):
        """
        TDD测试：UnsupportedFileTypeError应该列出支持的格式
        
        这个测试覆盖第68-76行的UnsupportedFileTypeError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        extension = ".xyz"
        supported = [".xlsx", ".xls", ".csv"]
        error = UnsupportedFileTypeError(extension, supported)
        
        assert extension in error.message
        assert "支持的格式" in error.message
        assert error.error_code == "UNSUPPORTED_FILE_TYPE"
        assert error.details["file_extension"] == extension
        assert error.details["supported_formats"] == supported

    def test_corrupted_file_error_default_reason(self):
        """
        TDD测试：CorruptedFileError应该使用默认原因
        
        这个测试覆盖第79-87行的CorruptedFileError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        file_path = "/path/to/corrupted.xlsx"
        error = CorruptedFileError(file_path)
        
        assert "文件损坏" in error.message
        assert "文件格式损坏" in error.message
        assert error.error_code == "CORRUPTED_FILE"
        assert error.details["reason"] == "文件格式损坏"

    def test_corrupted_file_error_custom_reason(self):
        """
        TDD测试：CorruptedFileError应该处理自定义原因
        
        这个测试确保自定义原因被正确处理
        """
        # 🔴 红阶段：编写测试描述期望的行为
        file_path = "/path/to/file.xlsx"
        custom_reason = "ZIP archive corrupted"
        error = CorruptedFileError(file_path, custom_reason)
        
        assert custom_reason in error.message
        assert error.details["reason"] == custom_reason

class TestParseErrors:
    """测试解析相关的异常类。"""

    def test_sheet_not_found_error_without_available_sheets(self):
        """
        TDD测试：SheetNotFoundError应该处理没有可用工作表列表的情况
        
        这个测试覆盖第95-104行的SheetNotFoundError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        sheet_name = "NonExistentSheet"
        error = SheetNotFoundError(sheet_name)
        
        assert sheet_name in error.message
        assert error.error_code == "SHEET_NOT_FOUND"
        assert error.details["available_sheets"] == []

    def test_sheet_not_found_error_with_available_sheets(self):
        """
        TDD测试：SheetNotFoundError应该列出可用的工作表
        
        这个测试确保可用工作表列表被正确包含
        """
        # 🔴 红阶段：编写测试描述期望的行为
        sheet_name = "Missing"
        available = ["Sheet1", "Sheet2", "Data"]
        error = SheetNotFoundError(sheet_name, available)
        
        assert "可用工作表" in error.message
        assert error.details["available_sheets"] == available

    def test_invalid_range_error_default_reason(self):
        """
        TDD测试：InvalidRangeError应该使用默认原因
        
        这个测试覆盖第107-115行的InvalidRangeError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        range_string = "INVALID_RANGE"
        error = InvalidRangeError(range_string)
        
        assert range_string in error.message
        assert "范围格式无效" in error.message
        assert error.error_code == "INVALID_RANGE"

    def test_invalid_range_error_custom_reason(self):
        """
        TDD测试：InvalidRangeError应该处理自定义原因
        
        这个测试确保自定义原因被正确处理
        """
        # 🔴 红阶段：编写测试描述期望的行为
        range_string = "A1:Z999999"
        custom_reason = "Range too large"
        error = InvalidRangeError(range_string, custom_reason)
        
        assert custom_reason in error.message
        assert error.details["reason"] == custom_reason

    def test_style_extraction_error(self):
        """
        TDD测试：StyleExtractionError应该包含单元格引用和原因
        
        这个测试覆盖第118-126行的StyleExtractionError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        cell_ref = "A1"
        reason = "Invalid color format"
        error = StyleExtractionError(cell_ref, reason)
        
        assert cell_ref in error.message
        assert reason in error.message
        assert error.error_code == "STYLE_EXTRACTION_ERROR"
        assert error.details["cell_reference"] == cell_ref
        assert error.details["reason"] == reason

class TestConversionErrors:
    """测试转换相关的异常类。"""

    def test_html_conversion_error_without_sheet_name(self):
        """
        TDD测试：HTMLConversionError应该处理没有工作表名的情况
        
        这个测试覆盖第134-142行的HTMLConversionError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        reason = "Template rendering failed"
        error = HTMLConversionError(reason)
        
        assert reason in error.message
        assert error.error_code == "HTML_CONVERSION_ERROR"
        assert error.details["sheet_name"] is None

    def test_html_conversion_error_with_sheet_name(self):
        """
        TDD测试：HTMLConversionError应该处理包含工作表名的情况
        
        这个测试确保工作表名被正确包含
        """
        # 🔴 红阶段：编写测试描述期望的行为
        reason = "CSS generation failed"
        sheet_name = "DataSheet"
        error = HTMLConversionError(reason, sheet_name)
        
        assert error.details["sheet_name"] == sheet_name

    def test_validation_error(self):
        """
        TDD测试：ValidationError应该包含字段、值和原因
        
        这个测试覆盖第145-153行的ValidationError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        field = "email"
        value = "invalid-email"
        reason = "Invalid email format"
        error = ValidationError(field, value, reason)
        
        assert field in error.message
        assert str(value) in error.message
        assert reason in error.message
        assert error.error_code == "VALIDATION_ERROR"
        assert error.details["field"] == field
        assert error.details["value"] == value
        assert error.details["reason"] == reason

class TestResourceErrors:
    """测试资源相关的异常类。"""

    def test_memory_limit_exceeded_error(self):
        """
        TDD测试：MemoryLimitExceededError应该包含内存使用信息

        这个测试覆盖第161-169行的MemoryLimitExceededError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        current_usage = 512
        limit = 256
        error = MemoryLimitExceededError(current_usage, limit)

        assert str(current_usage) in error.message
        assert str(limit) in error.message
        assert "内存使用超出限制" in error.message
        assert error.error_code == "MEMORY_LIMIT_EXCEEDED"
        assert error.details["current_usage"] == current_usage
        assert error.details["limit"] == limit

    def test_file_size_limit_exceeded_error(self):
        """
        TDD测试：FileSizeLimitExceededError应该包含文件大小信息

        这个测试覆盖第172-180行的FileSizeLimitExceededError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        file_size = 100
        limit = 50
        file_path = "/path/to/large/file.xlsx"
        error = FileSizeLimitExceededError(file_size, limit, file_path)

        assert str(file_size) in error.message
        assert str(limit) in error.message
        assert file_path in error.message
        assert error.error_code == "FILE_SIZE_LIMIT_EXCEEDED"
        assert error.details["file_size"] == file_size
        assert error.details["limit"] == limit
        assert error.details["file_path"] == file_path

    def test_timeout_error(self):
        """
        TDD测试：TimeoutError应该包含操作和超时信息

        这个测试覆盖第183-191行的TimeoutError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        operation = "file parsing"
        timeout_seconds = 30
        error = TimeoutError(operation, timeout_seconds)

        assert operation in error.message
        assert str(timeout_seconds) in error.message
        assert "操作超时" in error.message
        assert error.error_code == "OPERATION_TIMEOUT"
        assert error.details["operation"] == operation
        assert error.details["timeout_seconds"] == timeout_seconds

class TestOtherErrors:
    """测试其他异常类。"""

    def test_cache_error(self):
        """
        TDD测试：CacheError应该包含操作和原因信息

        这个测试覆盖第194-202行的CacheError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        operation = "cache_get"
        reason = "Redis connection failed"
        error = CacheError(operation, reason)

        assert operation in error.message
        assert reason in error.message
        assert "缓存操作失败" in error.message
        assert error.error_code == "CACHE_ERROR"
        assert error.details["operation"] == operation
        assert error.details["reason"] == reason

    def test_configuration_error(self):
        """
        TDD测试：ConfigurationError应该包含配置键和原因

        这个测试覆盖第205-213行的ConfigurationError代码路径
        """
        # 🔴 红阶段：编写测试描述期望的行为
        config_key = "database.host"
        reason = "Invalid hostname format"
        error = ConfigurationError(config_key, reason)

        assert config_key in error.message
        assert reason in error.message
        assert "配置错误" in error.message
        assert error.error_code == "CONFIGURATION_ERROR"
        assert error.details["config_key"] == config_key
        assert error.details["reason"] == reason

class TestErrorCodeMap:
    """测试错误代码映射功能。"""

    def test_error_code_map_completeness(self):
        """
        TDD测试：ERROR_CODE_MAP应该包含所有定义的错误代码

        这个测试覆盖第217-232行的ERROR_CODE_MAP定义
        """
        # 🔴 红阶段：编写测试描述期望的行为
        expected_codes = [
            "FILE_NOT_FOUND", "FILE_ACCESS_DENIED", "UNSUPPORTED_FILE_TYPE",
            "CORRUPTED_FILE", "SHEET_NOT_FOUND", "INVALID_RANGE",
            "STYLE_EXTRACTION_ERROR", "HTML_CONVERSION_ERROR", "VALIDATION_ERROR",
            "MEMORY_LIMIT_EXCEEDED", "FILE_SIZE_LIMIT_EXCEEDED", "OPERATION_TIMEOUT",
            "CACHE_ERROR", "CONFIGURATION_ERROR"
        ]

        for code in expected_codes:
            assert code in ERROR_CODE_MAP
            assert issubclass(ERROR_CODE_MAP[code], SheetParserError)

    def test_get_error_by_code_existing_code(self):
        """
        TDD测试：get_error_by_code应该返回正确的异常类

        这个测试覆盖第235-237行的get_error_by_code函数
        """
        # 🔴 红阶段：编写测试描述期望的行为
        error_class = get_error_by_code("FILE_NOT_FOUND")
        assert error_class == FileNotFoundError

    def test_get_error_by_code_unknown_code(self):
        """
        TDD测试：get_error_by_code应该为未知代码返回基础异常类

        这个测试确保未知错误代码被正确处理
        """
        # 🔴 红阶段：编写测试描述期望的行为
        error_class = get_error_by_code("UNKNOWN_ERROR_CODE")
        assert error_class == SheetParserError

class TestInheritanceHierarchy:
    """测试异常类的继承层次结构。"""

    def test_file_error_inheritance(self):
        """
        TDD测试：所有文件错误应该继承自FileError

        这个测试验证异常类的继承关系
        """
        # 🔴 红阶段：编写测试描述期望的行为
        assert issubclass(FileError, SheetParserError)
        assert issubclass(FileNotFoundError, FileError)
        assert issubclass(FileAccessError, FileError)
        assert issubclass(UnsupportedFileTypeError, FileError)
        assert issubclass(CorruptedFileError, FileError)

    def test_parse_error_inheritance(self):
        """
        TDD测试：所有解析错误应该继承自ParseError

        这个测试验证解析错误的继承关系
        """
        # 🔴 红阶段：编写测试描述期望的行为
        assert issubclass(ParseError, SheetParserError)
        assert issubclass(SheetNotFoundError, ParseError)
        assert issubclass(InvalidRangeError, ParseError)
        assert issubclass(StyleExtractionError, ParseError)

    def test_conversion_error_inheritance(self):
        """
        TDD测试：所有转换错误应该继承自ConversionError

        这个测试验证转换错误的继承关系
        """
        # 🔴 红阶段：编写测试描述期望的行为
        assert issubclass(ConversionError, SheetParserError)
        assert issubclass(HTMLConversionError, ConversionError)

    def test_resource_error_inheritance(self):
        """
        TDD测试：所有资源错误应该继承自ResourceError

        这个测试验证资源错误的继承关系
        """
        # 🔴 红阶段：编写测试描述期望的行为
        assert issubclass(ResourceError, SheetParserError)
        assert issubclass(MemoryLimitExceededError, ResourceError)
        assert issubclass(FileSizeLimitExceededError, ResourceError)
        assert issubclass(TimeoutError, ResourceError)

    def test_all_errors_inherit_from_base(self):
        """
        TDD测试：所有自定义异常都应该继承自SheetParserError

        这个测试确保所有异常都有正确的基类
        """
        # 🔴 红阶段：编写测试描述期望的行为
        all_error_classes = [
            FileError, FileNotFoundError, FileAccessError, UnsupportedFileTypeError,
            CorruptedFileError, ParseError, SheetNotFoundError, InvalidRangeError,
            StyleExtractionError, ConversionError, HTMLConversionError, ValidationError,
            ResourceError, MemoryLimitExceededError, FileSizeLimitExceededError,
            TimeoutError, CacheError, ConfigurationError
        ]

        for error_class in all_error_classes:
            assert issubclass(error_class, SheetParserError)
