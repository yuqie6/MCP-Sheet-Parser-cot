from dataclasses import dataclass, field
from typing import Any, List, Optional

@dataclass
class Style:
    bold: bool = False
    italic: bool = False
    font_color: str = "#000000"
    background_color: str = "#FFFFFF"

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