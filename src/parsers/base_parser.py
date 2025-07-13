"""
基础解析器模块

定义所有解析器的抽象基类。
"""

from abc import ABC, abstractmethod
from typing import Any, Dict, Optional
from src.models.table_model import Sheet, LazySheet


class BaseParser(ABC):
    """
    Abstract base class for all file parsers.

    Defines the common interface that all parsers must implement. This includes
    methods for parsing a file into a Sheet object, checking for streaming
    support, and creating a lazy-loading sheet.
    """

    @abstractmethod
    def parse(self, file_path: str) -> Sheet:
        """
        Parses the given file and returns a Sheet object.

        This is the primary method for each parser. It should handle opening the
        file, reading its content and structure, and converting it into a
        standardized Sheet object.

        Args:
            file_path: The absolute path to the file to be parsed.

        Returns:
            A Sheet object containing the structured data and styles.

        Raises:
            RuntimeError: If parsing fails for any reason.
        """
        pass
    
    def supports_streaming(self) -> bool:
        """Check if this parser supports streaming."""
        return False
    
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> Optional[LazySheet]:
        """
        Creates a LazySheet for streaming data on demand.

        This method should be implemented by parsers that support streaming.
        It returns a LazySheet object which can load data in chunks instead of
        all at once.

        Args:
            file_path: The absolute path to the file.
            sheet_name: The name of the sheet to parse (optional).

        Returns:
            A LazySheet object if streaming is supported, otherwise None.
        """
    def _style_to_dict(self, style: Any) -> Dict[str, Any]:
        """
        将特定于库的Style对象转换为标准化的字典格式。
        这是一个辅助方法，可以在子类中实现，以供CoreService使用。
        """
        if not style:
            return {}
        # 默认实现返回一个空字典，子类应重写此方法
        return {}

