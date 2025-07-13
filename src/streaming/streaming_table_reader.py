"""
流式表格读取器模块

提供内存高效的流式 API 层，封装任何解析器并提供
统一的接口来分块读取大文件，支持可选过滤。
"""

from typing import Iterator, Optional, List, Dict, Any, Tuple
from pathlib import Path
from dataclasses import dataclass
import re
import logging

from ..models.table_model import Sheet, Row, Cell, LazySheet
from ..parsers.factory import ParserFactory
from ..parsers.base_parser import BaseParser

logger = logging.getLogger(__name__)


@dataclass
class ChunkFilter:
    """流式读取时的数据过滤配置。"""
    columns: Optional[List[str]] = None  # 要包含的列名
    column_indices: Optional[List[int]] = None  # 要包含的列索引（替代列名）
    start_row: int = 0  # 起始行索引（基于0）
    max_rows: Optional[int] = None  # 最大读取行数
    range_string: Optional[str] = None  # Excel样式的范围，如"A1:D10"


@dataclass
class StreamingChunk:
    """表示流式操作中的数据块。"""
    rows: List[Row]  # 数据行
    headers: List[str]  # 表头
    chunk_index: int  # 块索引
    total_chunks: Optional[int] = None  # 总块数
    start_row: int = 0  # 起始行
    metadata: Optional[Dict[str, Any]] = None  # 元数据


class StreamingTableReader:
    """
    内存高效的流式表格读取器，封装任何解析器。
    
    提供统一的接口来分块读取大文件，支持列子集选择
    和范围过滤器进行早期数据剪枝。
    """
    
    def __init__(self, file_path: str, parser: Optional[BaseParser] = None):
        """
        初始化流式读取器。
        
        Args:
            file_path: 要读取的文件路径
            parser: 可选的特定解析器（如果为None则自动检测）
        """
        self.file_path = Path(file_path)
        self._parser = parser or ParserFactory.get_parser(str(file_path))
        self._lazy_sheet: Optional[LazySheet] = None
        self._regular_sheet: Optional[Sheet] = None
        self._total_rows_cache: Optional[int] = None
        self._headers_cache: Optional[List[str]] = None
        
        # 初始化数据源
        self._init_data_source()
    
    def _init_data_source(self):
        """初始化适当的数据源（懒加载或常规）。"""
        if self._parser.supports_streaming():
            self._lazy_sheet = self._parser.create_lazy_sheet(str(self.file_path))
            logger.info(f"为 {self.file_path} 初始化了懒加载工作表")
        else:
            self._regular_sheet = self._parser.parse(str(self.file_path))
            logger.info(f"为 {self.file_path} 初始化了常规工作表")
    
    def iter_chunks(self, rows: int = 1000, filter_config: Optional[ChunkFilter] = None) -> Iterator[StreamingChunk]:
        """
        分块迭代数据。
        
        Args:
            rows: 每块的行数
            filter_config: 可选的数据过滤配置
            
        Yields:
            包含过滤数据的StreamingChunk对象
        """
        if rows <= 0:
            raise ValueError("块大小必须为正数")
        
        # 如果指定了范围过滤器，应用它
        if filter_config and filter_config.range_string:
            range_filter = self._parse_range_filter(filter_config.range_string)
            start_row = range_filter['start_row']
            max_rows = range_filter['max_rows']
            column_indices = range_filter['column_indices']
        else:
            start_row = filter_config.start_row if filter_config else 0
            max_rows = filter_config.max_rows if filter_config else None
            column_indices = None
        
        # 获取表头
        headers = self._get_headers()
        
        # 对表头应用列过滤
        if filter_config:
            headers, column_indices = self._apply_column_filter(headers, filter_config, column_indices)
        
        # 计算总块数
        total_rows = self._get_total_rows()
        if max_rows:
            total_rows = min(total_rows - start_row, max_rows)
        else:
            total_rows = total_rows - start_row
        
        total_chunks = (total_rows + rows - 1) // rows if total_rows > 0 else 0
        
        # 迭代块
        chunk_index = 0
        current_row = start_row
        
        while current_row < self._get_total_rows():
            # 计算块大小
            chunk_size = min(rows, self._get_total_rows() - current_row)
            if max_rows:
                chunk_size = min(chunk_size, max_rows - (current_row - start_row))
            
            if chunk_size <= 0:
                break
            
            # 获取块数据
            chunk_rows = self._get_chunk_rows(current_row, chunk_size, column_indices)
            
            # 创建块
            chunk = StreamingChunk(
                rows=chunk_rows,
                headers=headers,
                chunk_index=chunk_index,
                total_chunks=total_chunks,
                start_row=current_row,
                metadata={
                    'file_path': str(self.file_path),
                    'chunk_size': len(chunk_rows),
                    'filtered_columns': len(headers),
                    'total_columns': len(self._get_headers()),
                    'supports_streaming': self._parser.supports_streaming()
                }
            )
            
            yield chunk
            
            current_row += chunk_size
            chunk_index += 1
    
    def _get_headers(self) -> List[str]:
        """从第一行获取表头。"""
        if self._headers_cache is None:
            if self._lazy_sheet:
                first_row = self._lazy_sheet.get_row(0)
            else:
                first_row = self._regular_sheet.rows[0]
            
            self._headers_cache = [
                cell.value if cell.value is not None else f"Column_{i}"
                for i, cell in enumerate(first_row.cells)
            ]
        
        return self._headers_cache
    
    def _get_total_rows(self) -> int:
        """获取总行数。"""
        if self._total_rows_cache is None:
            if self._lazy_sheet:
                self._total_rows_cache = self._lazy_sheet.get_total_rows()
            else:
                self._total_rows_cache = len(self._regular_sheet.rows)
        
        return self._total_rows_cache
    
    def _get_chunk_rows(self, start_row: int, chunk_size: int, column_indices: Optional[List[int]] = None) -> List[Row]:
        """获取一块数据行，支持可选的列过滤。"""
        if self._lazy_sheet:
            # 使用懒加载工作表进行流式读取
            rows = list(self._lazy_sheet.iter_rows(start_row, chunk_size))
        else:
            # 使用常规工作表
            end_row = min(start_row + chunk_size, len(self._regular_sheet.rows))
            rows = self._regular_sheet.rows[start_row:end_row]
        
        # 应用列过滤
        if column_indices:
            filtered_rows = []
            for row in rows:
                filtered_cells = [
                    row.cells[i] if i < len(row.cells) else Cell(value=None)
                    for i in column_indices
                ]
                filtered_rows.append(Row(cells=filtered_cells))
            return filtered_rows
        
        return rows
    
    def _apply_column_filter(self, headers: List[str], filter_config: ChunkFilter, 
                           existing_indices: Optional[List[int]] = None) -> Tuple[List[str], List[int]]:
        """对表头应用列过滤并返回过滤后的表头和索引。"""
        if existing_indices:
            # 使用来自范围过滤器的现有索引
            filtered_headers = [headers[i] for i in existing_indices if i < len(headers)]
            return filtered_headers, existing_indices
        
        if filter_config.column_indices:
            # 按列索引过滤
            indices = [i for i in filter_config.column_indices if 0 <= i < len(headers)]
            filtered_headers = [headers[i] for i in indices]
            return filtered_headers, indices
        
        if filter_config.columns:
            # 按列名过滤
            indices = []
            filtered_headers = []
            for col_name in filter_config.columns:
                try:
                    idx = headers.index(col_name)
                    indices.append(idx)
                    filtered_headers.append(col_name)
                except ValueError:
                    logger.warning(f"在表头中未找到列 '{col_name}'")
            
            return filtered_headers, indices
        
        # 不过滤
        return headers, list(range(len(headers)))
    
    def _parse_range_filter(self, range_string: str) -> Dict[str, Any]:
        """解析Excel样式的范围字符串，如'A1:D10'。"""
        range_string = range_string.strip().upper()
        
        # A1:D10类型的范围模式
        range_pattern = r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$'
        match = re.match(range_pattern, range_string)
        
        if not match:
            raise ValueError(f"无效的范围格式: {range_string}。期望格式: A1:D10")
        
        start_col_str, start_row_str, end_col_str, end_row_str = match.groups()
        
        # 将列字母转换为索引
        start_col = self._col_to_index(start_col_str)
        end_col = self._col_to_index(end_col_str)
        
        # 将行号转换为基于0的索引
        start_row = int(start_row_str) - 1
        end_row = int(end_row_str) - 1
        
        return {
            'start_row': start_row,
            'max_rows': end_row - start_row + 1,
            'column_indices': list(range(start_col, end_col + 1))
        }
    
    def _col_to_index(self, col_str: str) -> int:
        """将列字母转换为基于0的索引 (A=0, B=1, ..., Z=25, AA=26, ...)。"""
        result = 0
        for char in col_str:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1
    
    def get_info(self) -> Dict[str, Any]:
        """获取文件和读取器的信息。"""
        return {
            'file_path': str(self.file_path),
            'file_size': self.file_path.stat().st_size,
            'parser_type': type(self._parser).__name__,
            'supports_streaming': self._parser.supports_streaming(),
            'total_rows': self._get_total_rows(),
            'total_columns': len(self._get_headers()),
            'headers': self._get_headers(),
            'estimated_memory_usage': self._estimate_memory_usage()
        }
    
    def _estimate_memory_usage(self) -> str:
        """估算文件的内存使用量。"""
        total_rows = self._get_total_rows()
        total_cols = len(self._get_headers())
        
        # 粗略估算：假设每个单元格平均咇50字节
        estimated_bytes = total_rows * total_cols * 50
        
        if estimated_bytes < 1024:
            return f"{estimated_bytes} bytes"
        elif estimated_bytes < 1024 * 1024:
            return f"{estimated_bytes / 1024:.1f} KB"
        elif estimated_bytes < 1024 * 1024 * 1024:
            return f"{estimated_bytes / (1024 * 1024):.1f} MB"
        else:
            return f"{estimated_bytes / (1024 * 1024 * 1024):.1f} GB"
    
    def __enter__(self):
        """上下文管理器入口。"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器退出 - 清理资源。"""
        # 如果需要，清理任何资源
        pass
