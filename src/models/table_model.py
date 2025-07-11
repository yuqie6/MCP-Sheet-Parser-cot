from dataclasses import dataclass, field
from typing import List, Optional, Any

@dataclass
class Style:
    """Represents the styling of a cell."""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    font_color: Optional[str] = None
    fill_color: Optional[str] = None

@dataclass
class Cell:
    """Represents a single cell in a sheet."""
    value: Any
    row: int
    column: int
    style: Optional[Style] = None

@dataclass
class Row:
    """Represents a single row in a sheet."""
    cells: List[Cell] = field(default_factory=list)

@dataclass
class Sheet:
    """Represents a single sheet in a spreadsheet."""
    name: Optional[str] = None
    rows: List[Row] = field(default_factory=list)
    merged_cells: List[str] = field(default_factory=list)


@dataclass
class Workbook:
    """Represents a workbook that can contain multiple sheets."""
    sheets: List[Sheet] = field(default_factory=list)
