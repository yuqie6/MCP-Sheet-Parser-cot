from dataclasses import dataclass, field
from typing import Any, Iterator, Optional, Protocol, Union
from abc import ABC, abstractmethod

@dataclass
# 富文本片段的样式定义
class RichTextFragmentStyle:
    """表示富文本片段的样式。"""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_color: Optional[str] = None
    font_size: Optional[float] = None
    font_name: Optional[str] = None

@dataclass
# 富文本片段定义
class RichTextFragment:
    """表示带有独立样式的文本片段。"""
    text: str
    style: RichTextFragmentStyle

# 单元格的值可以是简单类型或富文本片段列表
CellValue = Union[Any, list[RichTextFragment]]

class LazyRowProvider(Protocol):
    """惰性行提供者协议，可按需流式获取行。"""

    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator['Row']:
        """从 start_row 开始按需生成行。"""
        ...

    def get_row(self, row_index: int) -> 'Row':
        """按索引获取指定行。"""
        ...

    def get_total_rows(self) -> int:
        """无需加载全部数据即可获取总行数。"""
        ...

class StreamingCapable(ABC):
    """支持流式处理的解析器基类。"""

    @abstractmethod
    def supports_streaming(self) -> bool:
        """检查该解析器是否支持流式处理。"""
        return False

    @abstractmethod
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> 'LazySheet':
        """创建可按需流式读取数据的惰性表对象。"""
        pass


@dataclass
class Style:
    """
    单元格的视觉样式。

    此类包含所有样式信息，包括字体属性、颜色、对齐方式、边框和其他格式细节。
    提供跨不同文件格式统一处理样式的方式。
    """
    # 字体属性（富文本后将弃用）
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_color: Optional[str] = None
    font_size: Optional[float] = None
    font_name: Optional[str] = None

    # 背景与填充
    background_color: Optional[str] = None

    # 文本对齐
    text_align: Optional[str] = None  # 可选：left, center, right, justify
    vertical_align: Optional[str] = None  # 可选：top, middle, bottom

    # 边框属性
    border_top: Optional[str] = None
    border_bottom: Optional[str] = None
    border_left: Optional[str] = None
    border_right: Optional[str] = None
    border_color: Optional[str] = None

    # 自动换行与格式
    wrap_text: bool = False
    number_format: Optional[str] = None

    # 高级特性
    formula: Optional[str] = None     # 单元格公式字符串
    hyperlink: Optional[str] = None  # 超链接 URL
    comment: Optional[str] = None    # 单元格批注



@dataclass
class Cell:
    """
    表格中的单个单元格。

    单元格包含其值、可选的样式对象，以及合并单元格的行/列跨度信息。
    """
    value: CellValue
    style: Style | None = None
    row_span: int = 1
    col_span: int = 1
    formula: Optional[str] = None




@dataclass
class Row:
    """
    表格中的一行，包含若干单元格对象。
    """
    cells: list[Cell]



@dataclass
class ChartPosition:
    """图表的精确定位信息。"""
    from_col: int  # 起始列索引（0基）
    from_row: int  # 起始行索引（0基）
    from_col_offset: float  # 起始列内偏移（像素）
    from_row_offset: float  # 起始行内偏移（像素）
    to_col: int  # 结束列索引（0基）
    to_row: int  # 结束行索引（0基）
    to_col_offset: float  # 结束列内偏移（像素）
    to_row_offset: float  # 结束行内偏移（像素）
    
    def get_width_in_columns(self) -> int:
        """获取图表占用的列数。"""
        return self.to_col - self.from_col + 1
    
    def get_height_in_rows(self) -> int:
        """获取图表占用的行数。"""
        return self.to_row - self.from_row + 1

@dataclass
class Chart:
    """表格中的图表对象。"""
    name: str
    type: str  # 如 'bar', 'line', 'pie'
    anchor: Optional[str] = None  # 锚点位置，如 'A1'
    
    # 原始图表数据，用于SVG渲染
    chart_data: Optional[dict] = None  # 原始图表数据，包含系列、标题等
    svg_data: Optional[str] = None  # 渲染后的SVG字符串
    
    # 精确定位信息
    position: Optional[ChartPosition] = None  # 详细的定位和尺寸信息



@dataclass
class Sheet:
    """
    表格的完整工作表对象，包含名称、所有行和合并单元格信息。
    此类将所有数据存储于内存中。
    """
    name: str
    rows: list[Row]
    merged_cells: list[str] = field(default_factory=list)
    charts: list[Chart] = field(default_factory=list)
    column_widths: dict[int, float] = field(default_factory=dict)  # 列宽信息 {列索引: 宽度}
    row_heights: dict[int, float] = field(default_factory=dict)    # 行高信息 {行索引: 高度}
    default_column_width: float = 8.43  # Excel默认列宽
    default_row_height: float = 18.0    # Excel默认行高

    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator[Row]:
        """
        遍历表格中的部分行。

        参数：
            start_row: 起始行索引。
            max_rows: 遍历的最大行数。

        产出：
            指定范围内的 Row 对象。
        """
        end_row = len(self.rows)
        if max_rows is not None:
            end_row = min(start_row + max_rows, len(self.rows))

        for i in range(start_row, end_row):
            if i < len(self.rows):
                yield self.rows[i]

    def get_total_rows(self) -> int:
        """返回表格的总行数。"""
        return len(self.rows)


class LazySheet:
    """惰性表对象，可按需流式读取数据，无需一次性加载全部内容到内存。"""

    def __init__(self, name: str, provider: LazyRowProvider, merged_cells: Optional[list[str]] = None):
        self.name = name
        self._provider = provider
        self.merged_cells = merged_cells or []
        self._total_rows_cache: Optional[int] = None

    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator[Row]:
        """按需遍历行。"""
        return self._provider.iter_rows(start_row, max_rows)

    def get_row(self, row_index: int) -> Row:
        """按索引获取指定行。"""
        return self._provider.get_row(row_index)

    def get_total_rows(self) -> int:
        """无需加载全部数据即可获取总行数。"""
        if self._total_rows_cache is None:
            self._total_rows_cache = self._provider.get_total_rows()
        return self._total_rows_cache

    def __getitem__(self, key) -> Row | list[Row]:
        """支持下标访问行。"""
        if isinstance(key, int):
            return self.get_row(key)
        elif isinstance(key, slice):
            start, stop, step = key.indices(self.get_total_rows())
            if step != 1:
                raise ValueError("暂不支持步长切片")
            return list(self.iter_rows(start, stop - start))
        else:
            raise TypeError("无效的下标类型")

    def to_sheet(self) -> Sheet:
        """将惰性表全部加载为常规 Sheet 对象。"""
        rows = list(self.iter_rows())
        return Sheet(
            name=self.name, 
            rows=rows, 
            merged_cells=self.merged_cells
        )
