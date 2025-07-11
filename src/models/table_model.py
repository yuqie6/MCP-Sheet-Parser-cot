from dataclasses import dataclass, field
from typing import Any, List, Optional

@dataclass
class Style:
    # Font properties
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_color: str = "#000000"
    font_size: float | None = None
    font_name: str | None = None

    # Background and fill
    background_color: str = "#FFFFFF"

    # Text alignment
    text_align: str = "left"  # left, center, right, justify
    vertical_align: str = "top"  # top, middle, bottom

    # Border properties
    border_top: str = ""
    border_bottom: str = ""
    border_left: str = ""
    border_right: str = ""
    border_color: str = "#000000"

    # Text wrapping and formatting
    wrap_text: bool = False
    number_format: str = ""

@dataclass
class Cell:
    value: Any
    style: Optional[Style] = None
    row_span: int = 1
    col_span: int = 1

@dataclass
class Row:
    cells: List[Cell]

@dataclass
class Sheet:
    name: str
    rows: List[Row]
    merged_cells: List[str] = field(default_factory=list)