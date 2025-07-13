from dataclasses import dataclass, field
from typing import Any, Iterator, Optional, Protocol
from abc import ABC, abstractmethod

class LazyRowProvider(Protocol):
    """Protocol for lazy row providers that can stream rows on demand."""
    
    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator['Row']:
        """Yield rows on demand starting from start_row."""
        ...
    
    def get_row(self, row_index: int) -> 'Row':
        """Get a specific row by index."""
        ...
    
    def get_total_rows(self) -> int:
        """Get total number of rows without loading all data."""
        ...

class StreamingCapable(ABC):
    """Base class for parsers that support streaming."""
    
    @abstractmethod
    def supports_streaming(self) -> bool:
        """Check if this parser supports streaming."""
        return False
    
    @abstractmethod
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> 'LazySheet':
        """Create a lazy sheet that can stream data on demand."""
        pass

@dataclass
class Style:
    """样式类，定义单元格的视觉样式属性。"""
    # 字体属性
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_color: str = "#000000"
    font_size: float | None = None
    font_name: str | None = None

    # 背景和填充
    background_color: str = "#FFFFFF"

    # 文本对齐
    text_align: str = "left"  # left, center, right, justify
    vertical_align: str = "top"  # top, middle, bottom

    # 边框属性
    border_top: str = ""
    border_bottom: str = ""
    border_left: str = ""
    border_right: str = ""
    border_color: str = "#000000"

    # 文本换行和格式化
    wrap_text: bool = False
    number_format: str = ""

    # 进阶功能
    hyperlink: str | None = None  # 超链接URL
    comment: str | None = None    # 单元格注释

@dataclass
class Cell:
    """单元格类，包含值、样式和合并信息。"""
    value: Any
    style: Style | None = None
    row_span: int = 1
    col_span: int = 1

@dataclass
class Row:
    """行类，包含一行中的所有单元格。"""
    cells: list[Cell]

@dataclass
class Sheet:
    """工作表类，包含表格的完整结构和数据。"""
    name: str
    rows: list[Row]
    merged_cells: list[str] = field(default_factory=list)
    
    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator[Row]:
        """Iterate over rows in the sheet."""
        end_row = len(self.rows)
        if max_rows is not None:
            end_row = min(start_row + max_rows, len(self.rows))
        
        for i in range(start_row, end_row):
            if i < len(self.rows):
                yield self.rows[i]
    
    def get_total_rows(self) -> int:
        """Get total number of rows."""
        return len(self.rows)

class LazySheet:
    """Lazy sheet that can stream data on demand without loading everything into memory."""
    
    def __init__(self, name: str, provider: LazyRowProvider, merged_cells: Optional[list[str]] = None):
        self.name = name
        self._provider = provider
        self.merged_cells = merged_cells or []
        self._total_rows_cache: Optional[int] = None
    
    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator[Row]:
        """Iterate over rows on demand."""
        return self._provider.iter_rows(start_row, max_rows)
    
    def get_row(self, row_index: int) -> Row:
        """Get a specific row by index."""
        return self._provider.get_row(row_index)
    
    def get_total_rows(self) -> int:
        """Get total number of rows without loading all data."""
        if self._total_rows_cache is None:
            self._total_rows_cache = self._provider.get_total_rows()
        return self._total_rows_cache
    
    def __getitem__(self, key) -> Row:
        """Support bracket notation for row access."""
        if isinstance(key, int):
            return self.get_row(key)
        elif isinstance(key, slice):
            start, stop, step = key.indices(self.get_total_rows())
            if step != 1:
                raise ValueError("Step slicing not supported")
            return list(self.iter_rows(start, stop - start))
        else:
            raise TypeError("Invalid key type")
    
    def to_sheet(self) -> Sheet:
        """Convert lazy sheet to regular sheet by loading all data."""
        rows = list(self.iter_rows())
        return Sheet(name=self.name, rows=rows, merged_cells=self.merged_cells)
