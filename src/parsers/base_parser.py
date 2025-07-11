"""
基础解析器模块

定义所有解析器的抽象基类。
"""

from abc import ABC, abstractmethod
from src.models.table_model import Sheet


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