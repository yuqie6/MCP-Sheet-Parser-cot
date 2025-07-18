"""
统一的异常处理模块

定义了项目中使用的所有自定义异常类，提供统一的错误处理机制。
"""

from typing import Any


class SheetParserError(Exception):
    """
    表格解析器基础异常类。
    
    所有自定义异常都应该继承自这个基类。
    """
    
    def __init__(self, message: str, error_code: str | None = None, details: dict[str, Any] | None = None):
        """
        初始化异常。
        
        参数:
            message: 错误消息
            error_code: 错误代码，用于程序化处理
            details: 额外的错误详情
        """
        super().__init__(message)
        self.message = message
        self.error_code = error_code or self.__class__.__name__
        self.details = details or {}
    
    def to_dict(self) -> dict[str, Any]:
        """将异常转换为字典格式，便于序列化。"""
        return {
            "error_type": self.__class__.__name__,
            "error_code": self.error_code,
            "message": self.message,
            "details": self.details
        }


class FileError(SheetParserError):
    """文件相关错误的基类。"""
    pass


class FileNotFoundError(FileError):
    """文件不存在错误。"""
    
    def __init__(self, file_path: str):
        super().__init__(
            f"文件不存在: {file_path}",
            "FILE_NOT_FOUND",
            {"file_path": file_path}
        )


class FileAccessError(FileError):
    """文件访问权限错误。"""
    
    def __init__(self, file_path: str, operation: str = "read"):
        super().__init__(
            f"无法{operation}文件: {file_path}",
            "FILE_ACCESS_DENIED",
            {"file_path": file_path, "operation": operation}
        )


class UnsupportedFileTypeError(FileError):
    """不支持的文件类型错误。"""
    
    def __init__(self, file_extension: str, supported_formats: list[str]):
        super().__init__(
            f"不支持的文件格式: '{file_extension}'. 支持的格式: {', '.join(supported_formats)}",
            "UNSUPPORTED_FILE_TYPE",
            {"file_extension": file_extension, "supported_formats": supported_formats}
        )


class CorruptedFileError(FileError):
    """文件损坏错误。"""
    
    def __init__(self, file_path: str, reason: str = "文件格式损坏"):
        super().__init__(
            f"文件损坏: {file_path} - {reason}",
            "CORRUPTED_FILE",
            {"file_path": file_path, "reason": reason}
        )


class ParseError(SheetParserError):
    """解析相关错误的基类。"""
    pass


class SheetNotFoundError(ParseError):
    """工作表不存在错误。"""
    
    def __init__(self, sheet_name: str, available_sheets: list[str] | None = None):
        available = f"可用工作表: {', '.join(available_sheets)}" if available_sheets else ""
        super().__init__(
            f"工作表不存在: '{sheet_name}'. {available}",
            "SHEET_NOT_FOUND",
            {"sheet_name": sheet_name, "available_sheets": available_sheets or []}
        )


class InvalidRangeError(ParseError):
    """无效范围错误。"""
    
    def __init__(self, range_string: str, reason: str = "范围格式无效"):
        super().__init__(
            f"无效的单元格范围: '{range_string}' - {reason}",
            "INVALID_RANGE",
            {"range_string": range_string, "reason": reason}
        )


class StyleExtractionError(ParseError):
    """样式提取错误。"""
    
    def __init__(self, cell_reference: str, reason: str):
        super().__init__(
            f"样式提取失败 {cell_reference}: {reason}",
            "STYLE_EXTRACTION_ERROR",
            {"cell_reference": cell_reference, "reason": reason}
        )


class ConversionError(SheetParserError):
    """转换相关错误的基类。"""
    pass


class HTMLConversionError(ConversionError):
    """HTML转换错误。"""
    
    def __init__(self, reason: str, sheet_name: str | None = None):
        super().__init__(
            f"HTML转换失败: {reason}",
            "HTML_CONVERSION_ERROR",
            {"reason": reason, "sheet_name": sheet_name}
        )


class ValidationError(SheetParserError):
    """数据验证错误。"""
    
    def __init__(self, field: str, value: Any, reason: str):
        super().__init__(
            f"验证失败 {field}='{value}': {reason}",
            "VALIDATION_ERROR",
            {"field": field, "value": value, "reason": reason}
        )


class ResourceError(SheetParserError):
    """资源相关错误的基类。"""
    pass


class MemoryLimitExceededError(ResourceError):
    """内存限制超出错误。"""
    
    def __init__(self, current_usage: int, limit: int):
        super().__init__(
            f"内存使用超出限制: {current_usage}MB > {limit}MB",
            "MEMORY_LIMIT_EXCEEDED",
            {"current_usage": current_usage, "limit": limit}
        )


class FileSizeLimitExceededError(ResourceError):
    """文件大小限制超出错误。"""
    
    def __init__(self, file_size: int, limit: int, file_path: str):
        super().__init__(
            f"文件大小超出限制: {file_size}MB > {limit}MB ({file_path})",
            "FILE_SIZE_LIMIT_EXCEEDED",
            {"file_size": file_size, "limit": limit, "file_path": file_path}
        )


class TimeoutError(ResourceError):
    """超时错误。"""
    
    def __init__(self, operation: str, timeout_seconds: int):
        super().__init__(
            f"操作超时: {operation} (超时时间: {timeout_seconds}秒)",
            "OPERATION_TIMEOUT",
            {"operation": operation, "timeout_seconds": timeout_seconds}
        )


class CacheError(SheetParserError):
    """缓存相关错误。"""
    
    def __init__(self, operation: str, reason: str):
        super().__init__(
            f"缓存操作失败: {operation} - {reason}",
            "CACHE_ERROR",
            {"operation": operation, "reason": reason}
        )


class ConfigurationError(SheetParserError):
    """配置错误。"""
    
    def __init__(self, config_key: str, reason: str):
        super().__init__(
            f"配置错误 {config_key}: {reason}",
            "CONFIGURATION_ERROR",
            {"config_key": config_key, "reason": reason}
        )


# 错误代码映射，用于快速查找
ERROR_CODE_MAP = {
    "FILE_NOT_FOUND": FileNotFoundError,
    "FILE_ACCESS_DENIED": FileAccessError,
    "UNSUPPORTED_FILE_TYPE": UnsupportedFileTypeError,
    "CORRUPTED_FILE": CorruptedFileError,
    "SHEET_NOT_FOUND": SheetNotFoundError,
    "INVALID_RANGE": InvalidRangeError,
    "STYLE_EXTRACTION_ERROR": StyleExtractionError,
    "HTML_CONVERSION_ERROR": HTMLConversionError,
    "VALIDATION_ERROR": ValidationError,
    "MEMORY_LIMIT_EXCEEDED": MemoryLimitExceededError,
    "FILE_SIZE_LIMIT_EXCEEDED": FileSizeLimitExceededError,
    "OPERATION_TIMEOUT": TimeoutError,
    "CACHE_ERROR": CacheError,
    "CONFIGURATION_ERROR": ConfigurationError,
}


def get_error_by_code(error_code: str) -> type[SheetParserError]:
    """根据错误代码获取异常类。"""
    return ERROR_CODE_MAP.get(error_code, SheetParserError)
