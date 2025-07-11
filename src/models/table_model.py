from dataclasses import dataclass, field
from typing import Any

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