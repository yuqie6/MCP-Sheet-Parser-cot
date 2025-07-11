from dataclasses import dataclass, field
from typing import List, Optional, Any

@dataclass
class Style:
    """Represents the styling of a cell."""
    # Basic text formatting (existing attributes - preserved for compatibility)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    font_color: Optional[str] = None
    fill_color: Optional[str] = None

    # Extended font properties
    font_size: Optional[int] = None
    font_name: Optional[str] = None

    # Text alignment
    alignment: Optional[str] = None  # 'left', 'center', 'right', 'justify'
    vertical_alignment: Optional[str] = None  # 'top', 'middle', 'bottom'

    # Border styling
    border_style: Optional[str] = None  # CSS border style string
    border_color: Optional[str] = None
    border_width: Optional[str] = None

    # Additional text decoration
    text_decoration: Optional[str] = None  # 'none', 'underline', 'line-through', etc.

    # Cell dimensions and spacing
    width: Optional[str] = None
    height: Optional[str] = None
    padding: Optional[str] = None

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
