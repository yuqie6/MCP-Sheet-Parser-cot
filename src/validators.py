"""
输入验证模块

提供统一的输入验证功能，确保数据安全性和完整性。
"""

import os
import re
from pathlib import Path

from .exceptions import (
    ValidationError,
    FileNotFoundError as CustomFileNotFoundError,
    FileAccessError,
    UnsupportedFileTypeError
)


class FileValidator:
    """文件相关的验证器。"""
    
    # 支持的文件扩展名
    SUPPORTED_EXTENSIONS = {'.csv', '.xlsx', '.xls', '.xlsb', '.xlsm'}
    
    # 最大文件大小 (MB)
    MAX_FILE_SIZE_MB = 500
    
    # 危险的文件路径模式
    DANGEROUS_PATH_PATTERNS = [
        r'\.\./',  # 路径遍历
        r'\.\.\\',  # Windows路径遍历
        r'/etc/',   # Unix系统目录
        r'/proc/',  # Unix进程目录
        r'/sys/',   # Unix系统目录
        r'C:\\Windows\\',  # Windows系统目录
        r'C:\\Program Files\\',  # Windows程序目录
    ]
    
    @classmethod
    def validate_file_path(cls, file_path: str) -> Path:
        """
        验证文件路径的安全性和有效性。
        
        Args:
            file_path: 文件路径字符串
            
        Returns:
            验证后的Path对象
            
        Raises:
            ValidationError: 路径验证失败
            FileNotFoundError: 文件不存在
            FileAccessError: 文件访问权限不足
        """
        if not file_path or not isinstance(file_path, str):
            raise ValidationError("file_path", file_path, "文件路径不能为空且必须是字符串")
        
        # 检查危险路径模式
        for pattern in cls.DANGEROUS_PATH_PATTERNS:
            if re.search(pattern, file_path, re.IGNORECASE):
                raise ValidationError("file_path", file_path, f"检测到危险路径模式: {pattern}")
        
        # 转换为Path对象并规范化
        try:
            path = Path(file_path).resolve()
        except (OSError, ValueError) as e:
            raise ValidationError("file_path", file_path, f"路径格式无效: {e}")
        
        # 检查文件是否存在
        if not path.exists():
            raise CustomFileNotFoundError(str(path))
        
        # 检查是否为文件（不是目录）
        if not path.is_file():
            raise ValidationError("file_path", file_path, "路径必须指向一个文件，不能是目录")
        
        # 检查文件访问权限
        if not os.access(path, os.R_OK):
            raise FileAccessError(str(path), "read")
        
        return path
    
    @classmethod
    def validate_file_extension(cls, file_path: str) -> str:
        """
        验证文件扩展名是否支持。
        
        Args:
            file_path: 文件路径
            
        Returns:
            小写的文件扩展名（不含点）
            
        Raises:
            UnsupportedFileTypeError: 不支持的文件类型
        """
        path = Path(file_path)
        extension = path.suffix.lower()
        
        if extension not in cls.SUPPORTED_EXTENSIONS:
            supported_list = [ext[1:] for ext in cls.SUPPORTED_EXTENSIONS]  # 移除点号
            raise UnsupportedFileTypeError(extension[1:] if extension else "", supported_list)
        
        return extension[1:]  # 返回不含点的扩展名
    
    @classmethod
    def validate_file_size(cls, file_path: str) -> int:
        """
        验证文件大小是否在允许范围内。
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件大小（字节）
            
        Raises:
            ValidationError: 文件大小超出限制
        """
        path = Path(file_path)
        file_size_bytes = path.stat().st_size
        file_size_mb = file_size_bytes / (1024 * 1024)
        
        if file_size_mb > cls.MAX_FILE_SIZE_MB:
            raise ValidationError(
                "file_size", 
                f"{file_size_mb:.2f}MB", 
                f"文件大小超出限制 (最大: {cls.MAX_FILE_SIZE_MB}MB)"
            )
        
        return file_size_bytes


class RangeValidator:
    """单元格范围验证器。"""
    
    # Excel单元格范围的正则表达式
    CELL_RANGE_PATTERN = re.compile(
        r'^([A-Z]+\d+)(?::([A-Z]+\d+))?$',
        re.IGNORECASE
    )
    
    # 单个单元格的正则表达式
    SINGLE_CELL_PATTERN = re.compile(
        r'^[A-Z]+\d+$',
        re.IGNORECASE
    )
    
    @classmethod
    def validate_range_string(cls, range_string: str) -> tuple[str, str] | None:
        """
        验证单元格范围字符串的格式。
        
        Args:
            range_string: 范围字符串，如 "A1:B10" 或 "A1"
            
        Returns:
            (start_cell, end_cell) 元组，如果是单个单元格则end_cell为None
            
        Raises:
            ValidationError: 范围格式无效
        """
        if not range_string:
            return None
        
        if not isinstance(range_string, str):
            raise ValidationError("range_string", range_string, "范围必须是字符串")
        
        range_string = range_string.strip().upper()
        
        # 检查是否为单个单元格
        if cls.SINGLE_CELL_PATTERN.match(range_string):
            return (range_string, range_string)
        
        # 检查是否为范围
        match = cls.CELL_RANGE_PATTERN.match(range_string)
        if not match:
            raise ValidationError(
                "range_string", 
                range_string, 
                "范围格式无效，应为 'A1' 或 'A1:B10' 格式"
            )
        
        start_cell, end_cell = match.groups()
        return (start_cell, end_cell or start_cell)


class DataValidator:
    """数据验证器。"""
    
    @staticmethod
    def validate_sheet_name(sheet_name: str) -> str:
        """
        验证工作表名称。
        
        Args:
            sheet_name: 工作表名称
            
        Returns:
            验证后的工作表名称
            
        Raises:
            ValidationError: 工作表名称无效
        """
        if not sheet_name:
            raise ValidationError("sheet_name", sheet_name, "工作表名称不能为空")
        
        if not isinstance(sheet_name, str):
            raise ValidationError("sheet_name", sheet_name, "工作表名称必须是字符串")
        
        # Excel工作表名称限制
        if len(sheet_name) > 31:
            raise ValidationError("sheet_name", sheet_name, "工作表名称不能超过31个字符")
        
        # 禁用字符
        forbidden_chars = ['\\', '/', '?', '*', '[', ']', ':']
        for char in forbidden_chars:
            if char in sheet_name:
                raise ValidationError(
                    "sheet_name", 
                    sheet_name, 
                    f"工作表名称不能包含字符: {char}"
                )
        
        return sheet_name.strip()
    
    @staticmethod
    def validate_page_size(page_size: int) -> int:
        """
        验证分页大小。
        
        Args:
            page_size: 分页大小
            
        Returns:
            验证后的分页大小
            
        Raises:
            ValidationError: 分页大小无效
        """
        if not isinstance(page_size, int):
            raise ValidationError("page_size", page_size, "分页大小必须是整数")
        
        if page_size <= 0:
            raise ValidationError("page_size", page_size, "分页大小必须大于0")
        
        if page_size > 10000:
            raise ValidationError("page_size", page_size, "分页大小不能超过10000")
        
        return page_size
    
    @staticmethod
    def validate_page_number(page_number: int) -> int:
        """
        验证页码。
        
        Args:
            page_number: 页码
            
        Returns:
            验证后的页码
            
        Raises:
            ValidationError: 页码无效
        """
        if not isinstance(page_number, int):
            raise ValidationError("page_number", page_number, "页码必须是整数")
        
        if page_number <= 0:
            raise ValidationError("page_number", page_number, "页码必须大于0")
        
        return page_number
    
    @staticmethod
    def validate_output_path(output_path: str, create_dirs: bool = True) -> Path:
        """
        验证输出路径。
        
        Args:
            output_path: 输出路径
            create_dirs: 是否创建不存在的目录
            
        Returns:
            验证后的Path对象
            
        Raises:
            ValidationError: 输出路径无效
        """
        if not output_path or not isinstance(output_path, str):
            raise ValidationError("output_path", output_path, "输出路径不能为空且必须是字符串")
        
        try:
            path = Path(output_path).resolve()
        except (OSError, ValueError) as e:
            raise ValidationError("output_path", output_path, f"路径格式无效: {e}")
        
        # 检查父目录
        parent_dir = path.parent
        if not parent_dir.exists():
            if create_dirs:
                try:
                    parent_dir.mkdir(parents=True, exist_ok=True)
                except OSError as e:
                    raise ValidationError("output_path", output_path, f"无法创建目录: {e}")
            else:
                raise ValidationError("output_path", output_path, f"目录不存在: {parent_dir}")
        
        # 检查写入权限
        if not os.access(parent_dir, os.W_OK):
            raise FileAccessError(str(parent_dir), "write")
        
        return path


def validate_file_input(file_path: str) -> tuple[Path, str]:
    """
    综合验证文件输入。
    
    Args:
        file_path: 文件路径
        
    Returns:
        (validated_path, file_extension) 元组
        
    Raises:
        各种验证异常
    """
    # 验证路径安全性和存在性
    validated_path = FileValidator.validate_file_path(file_path)
    
    # 验证文件扩展名
    file_extension = FileValidator.validate_file_extension(str(validated_path))
    
    # 验证文件大小
    FileValidator.validate_file_size(str(validated_path))
    
    return validated_path, file_extension
