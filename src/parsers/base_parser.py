"""
基础解析器模块

定义所有解析器的抽象基类。
"""

from abc import ABC, abstractmethod
from src.models.table_model import Sheet, LazySheet


class BaseParser(ABC):
    """
    所有文件解析器的抽象基类。

    定义所有解析器必须实现的通用接口，包括：
    - 解析文件为 Sheet 对象列表的方法
    - 检查是否支持流式处理的方法
    - 创建惰性加载表的方法
    """

    @abstractmethod
    def parse(self, file_path: str) -> list[Sheet]:
        """
        解析指定文件并返回 Sheet 对象列表。

        这是每个解析器的主要方法。应负责打开文件、读取内容和结构，并转换为标准化的 Sheet 对象列表。
        对于单工作表文件（如CSV），返回包含一个Sheet的列表。
        对于多工作表文件（如Excel），返回包含所有工作表的列表。

        参数：
            file_path: 要解析的文件的绝对路径。

        返回：
            包含结构化数据和样式的 Sheet 对象列表。

        异常：
            RuntimeError: 解析失败时抛出。
        """
        pass
    
    def supports_streaming(self) -> bool:
        """检查该解析器是否支持流式处理。"""
        return False
    
    def create_lazy_sheet(self, file_path: str, sheet_name: str | None = None) -> LazySheet | None:
        """
        创建用于按需流式读取数据的 LazySheet。

        仅支持流式处理的解析器应实现此方法。
        返回一个 LazySheet 对象，可分块加载数据而非一次性全部加载。

        参数：
            file_path: 文件的绝对路径。
            sheet_name: 要解析的工作表名称（可选）。

        返回：
            如果支持流式处理，返回 LazySheet 对象，否则返回 None。
        """
        return None

