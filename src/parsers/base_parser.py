"""
基础解析器模块

定义所有解析器的抽象基类。
"""

from abc import ABC, abstractmethod
from typing import Optional
from src.models.table_model import Sheet, LazySheet, StreamingCapable


class BaseParser(ABC):
    """解析器抽象基类，定义解析器的通用接口。"""

    @abstractmethod
    def parse(self, file_path: str) -> Sheet:
        """
        解析给定的文件并返回Sheet对象。

        Args:
            file_path: 要解析的文件路径

        Returns:
            包含表格数据和样式的Sheet对象

        Raises:
            RuntimeError: 当解析失败时
        """
        pass
    
    def supports_streaming(self) -> bool:
        """Check if this parser supports streaming."""
        return False
    
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> Optional[LazySheet]:
        """
        Create a lazy sheet that can stream data on demand.
        
        Args:
            file_path: 要解析的文件路径
            sheet_name: 工作表名称（可选）
            
        Returns:
            LazySheet对象，如果不支持流式读取则返回None
        """
        return None
