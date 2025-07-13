"""
内存高效表格读取的流式模块。

该模块提供StreamingTableReader抽象，封装任何解析器
并提供统一的接口来分块读取大文件。
"""

from .streaming_table_reader import StreamingTableReader, ChunkFilter, StreamingChunk

__all__ = ['StreamingTableReader', 'ChunkFilter', 'StreamingChunk']
