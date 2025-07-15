import logging
import re
from typing import Dict, Any

from src.constants import StyleConstants
from src.font_manager import get_font_manager
from src.models.table_model import Sheet, Style

logger = logging.getLogger(__name__)

# Border style mapping constant
BORDER_STYLE_MAP = {
    "thin": ("1px", "solid"),
    "medium": ("2px", "solid"),
    "thick": ("3px", "solid"),
    "solid": ("1px", "solid"),
    "dashed": ("1px", "dashed"),
    "dotted": ("1px", "dotted"),
    "double": ("3px", "double"),
    "groove": ("2px", "groove"),
    "ridge": ("2px", "ridge"),
    "inset": ("2px", "inset"),
    "outset": ("2px", "outset"),
    "hair": ("1px", "solid"),
    "mediumdashed": ("2px", "dashed"),
    "dashdot": ("1px", "dashed"),
    "mediumdashdot": ("2px", "dashed"),
    "dashdotdot": ("1px", "dashed"),
    "mediumdashdotdot": ("2px", "dashed"),
    "slantdashdot": ("1px", "dashed")
}


class StyleConverter:
    """Handles all CSS generation logic."""

    def collect_styles(self, sheet: Sheet) -> dict[str, Style]:
        """
        Collects all unique styles from a sheet.

        Args:
            sheet: The Sheet object.

        Returns:
            A dictionary of styles, with style IDs as keys and Style objects as values.
        """
        styles = {}
        style_counter = 0

        for row in sheet.rows:
            for cell in row.cells:
                if cell.style:
                    style_key = self.get_style_key(cell.style)
                    if style_key not in styles:
                        style_counter += 1
                        styles[f"style_{style_counter}"] = cell.style
        return styles

    def get_style_key(self, style: Style) -> str:
        """
        Generates a unique identifier for a style.

        Args:
            style: The Style object.

        Returns:
            A unique string identifier for the style.
        """
        key_parts = []
        if style.font_name:
            key_parts.append(f"fn:{style.font_name}")
        if style.font_size:
            key_parts.append(f"fs:{style.font_size}")
        if style.font_color:
            key_parts.append(f"fc:{style.font_color}")
        if style.background_color:
            key_parts.append(f"bg:{style.background_color}")
        if style.bold:
            key_parts.append("bold")
        if style.italic:
            key_parts.append("italic")
        if style.underline:
            key_parts.append("underline")
        if style.text_align:
            key_parts.append(f"ta:{style.text_align}")
        if style.vertical_align:
            key_parts.append(f"va:{style.vertical_align}")
        if style.border_top:
            key_parts.append(f"bt:{style.border_top}")
        if style.border_bottom:
            key_parts.append(f"bb:{style.border_bottom}")
        if style.border_left:
            key_parts.append(f"bl:{style.border_left}")
        if style.border_right:
            key_parts.append(f"br:{style.border_right}")
        if style.border_color:
            key_parts.append(f"bc:{style.border_color}")
        if style.wrap_text:
            key_parts.append("wrap")
        if style.number_format:
            key_parts.append(f"nf:{style.number_format}")
        if style.formula:
            key_parts.append(f"formula:{style.formula}")
        if style.hyperlink:
            key_parts.append(f"link:{style.hyperlink}")
        if style.comment:
            key_parts.append(f"comment:{style.comment}")

        return "|".join(key_parts) if key_parts else "default"

    def generate_css(self, styles: dict[str, Style], sheet: Sheet | None = None, position_calculator=None) -> str:
        """
        Generates CSS styles.

        Args:
            styles: A dictionary of styles.
            sheet: The Sheet object, for dimension information.

        Returns:
            A CSS string.
        """
        css_rules = [
            """
        table {
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            table-layout: fixed;
            width: auto;
            margin: 20px 0;
        }
        th, td {
            padding: 8px 12px;
            text-align: left;
            min-width: 60px;
            max-width: none;
            word-wrap: break-word;
            overflow: visible;
            white-space: normal;
            vertical-align: middle;
        }
        th {
            background-color: #f8f9fa;
            font-weight: bold;
            height: 20px;
            text-align: center;
            font-size: 11pt;
        }
        td {
            height: 18px;
            font-size: 10pt;
        }
        .wrap-text {
            white-space: pre-wrap !important;
            height: auto !important;
            min-height: 18px !important;
        }
        body {
            background-color: #ffffff !important;
            font-family: Arial, sans-serif;
            margin: 20px;
            line-height: 1.1;
        }
        .table-container {
            overflow-x: auto;
            margin: 20px 0;
        }
        .formula-cell {
            background-color: #f0f8ff !important;
            font-style: italic;
        }
        .number-cell {
            text-align: right;
        }
        .date-cell {
            text-align: center;
        }
        """
        ]

        for style_id, style in styles.items():
            css_rule = f".{style_id} {{"
            if style.font_name:
                font_family = self._format_font_family(style.font_name)
                css_rule += f" font-family: {font_family};"
            if style.font_size:
                font_size = self._format_font_size(style.font_size)
                css_rule += f" font-size: {font_size};"
            if style.font_color:
                formatted_color = self._format_color(style.font_color, is_font_color=True)
                css_rule += f" color: {formatted_color} !important;"
            if style.bold:
                css_rule += " font-weight: bold;"
            if style.italic:
                css_rule += " font-style: italic;"
            if style.underline:
                css_rule += " text-decoration: underline;"
            if style.background_color:
                formatted_bg_color = self._format_color(style.background_color, is_font_color=False,
                                                        is_border_color=False)
                if formatted_bg_color:
                    css_rule += f" background-color: {formatted_bg_color};"
            if style.text_align:
                css_rule += f" text-align: {style.text_align};"
            if style.vertical_align:
                css_rule += f" vertical-align: {style.vertical_align};"

            border_styles = self._generate_border_css(style)
            if border_styles:
                css_rule += border_styles
            if style.wrap_text:
                css_rule += " white-space: pre-wrap; word-wrap: break-word;"
            if style.number_format:
                css_rule += f" /* number-format: {style.number_format} */"
            css_rule += " }"
            css_rules.append(css_rule)

        if sheet:
            css_rules.append(self._generate_dimension_css(sheet))

        css_rules.append(self._generate_chart_css())

        return "\n".join(css_rules)

    def _generate_dimension_css(self, sheet: Sheet) -> str:
        """
        Generates CSS for column widths and row heights.
        """
        css_rules = []
        excel_to_px = 8.43
        for col_idx, width in sheet.column_widths.items():
            if width > 0:
                width_px = width * excel_to_px
                css_rules.append(
                    f"table td:nth-child({col_idx + 1}), table th:nth-child({col_idx + 1}) {{ width: {width_px:.0f}px; min-width: {width_px:.0f}px; }}")
        for row_idx, height in sheet.row_heights.items():
            if height > 0:
                css_rules.append(f"table tr:nth-child({row_idx + 1}) {{ height: {height}pt; }}")
        return "\n".join(css_rules)

    def _generate_chart_css(self) -> str:
        """Generates CSS for charts."""
        return """
        .table-container {
            position: relative;
            overflow: visible;
            margin: 20px 0;
        }
        .chart-overlay {
            position: absolute;
            z-index: 10;
            pointer-events: auto;
            background: rgba(255, 255, 255, 0.95);
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            padding: 8px;
        }
        .chart-overlay:hover {
            z-index: 20;
            box-shadow: 0 4px 16px rgba(0,0,0,0.2);
        }
        .chart-container {
            margin: 20px 0;
            padding: 15px;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            background-color: #fafafa;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .chart-container h3 {
            margin: 0 0 15px 0;
            color: #333;
            font-size: 16px;
            font-weight: bold;
        }
        .chart-svg-wrapper {
            display: flex;
            justify-content: center;
            align-items: center;
            background-color: white;
            border-radius: 4px;
            padding: 10px;
            min-height: 300px;
        }
        .chart-svg-wrapper svg {
            max-width: 100%;
            height: auto;
        }
        .chart-error, .chart-placeholder {
            text-align: center;
            padding: 40px 20px;
            color: #666;
            font-style: italic;
            background-color: #f5f5f5;
            border-radius: 4px;
            border: 1px dashed #ccc;
        }
        .chart-error {
            color: #d32f2f;
            background-color: #ffeaea;
            border-color: #ffcdd2;
        }
        @media (max-width: 768px) {
            .chart-container {
                margin: 10px 0;
                padding: 10px;
            }
            .chart-svg-wrapper {
                padding: 5px;
                min-height: 200px;
            }
            .chart-overlay {
                position: static !important;
                margin: 10px 0;
                z-index: auto;
            }
        }
        """

    def _generate_border_css(self, style: Style) -> str:
        """
        Generates CSS for borders.
        """
        border_css = ""
        has_any_border = False
        if style.border_top:
            border_css += f" border-top: {style.border_top} !important;"
            has_any_border = True
        if style.border_bottom:
            border_css += f" border-bottom: {style.border_bottom} !important;"
            has_any_border = True
        if style.border_left:
            border_css += f" border-left: {style.border_left} !important;"
            has_any_border = True
        if style.border_right:
            border_css += f" border-right: {style.border_right} !important;"
            has_any_border = True
        if not has_any_border:
            border_css = " border: none !important;"
        return border_css

    def _parse_border_style_complete(self, border_style: str, border_color: str) -> str:
        """
        Parses a complete border style string.
        """
        if not border_style:
            return f"1px solid {border_color}"
        border_lower = border_style.lower()
        if border_lower in BORDER_STYLE_MAP:
            width, style_type = BORDER_STYLE_MAP[border_lower]
            return f"{width} {style_type} {border_color}"
        pattern = r'(\d+(?:\.\d+)?)(px|pt|em|rem)?\s*(solid|dashed|dotted|double|groove|ridge|inset|outset)?'
        match = re.search(pattern, border_style.lower())
        if match:
            width = match.group(1)
            unit = match.group(2) or "px"
            style_type = match.group(3) or "solid"
            return f"{width}{unit} {style_type} {border_color}"
        return f"1px solid {border_color}"

    def _format_font_family(self, font_name: str) -> str:
        """
        Formats a font family name.
        """
        font_manager = get_font_manager()
        return font_manager.generate_font_family(font_name)

    def _format_font_size(self, font_size: float) -> str:
        """
        Formats a font size.
        """
        if not font_size or font_size <= 0:
            return f"{StyleConstants.DEFAULT_FONT_SIZE_PT}pt"
        if font_size < StyleConstants.MIN_FONT_SIZE_PT:
            adjusted_size = max(StyleConstants.MIN_FONT_SIZE_PT, font_size)
        elif font_size > StyleConstants.MAX_FONT_SIZE_PT:
            adjusted_size = min(StyleConstants.MAX_FONT_SIZE_PT, font_size)
        else:
            adjusted_size = font_size
        pt_size = round(adjusted_size, 1)
        if pt_size == int(pt_size):
            return f"{int(pt_size)}pt"
        else:
            return f"{pt_size}pt"

    def _format_color(self, color: str, is_font_color: bool = False, is_border_color: bool = False) -> str | None:
        """
        Formats a color string.
        """
        if not color:
            return StyleConstants.DEFAULT_FONT_COLOR if is_font_color else None
        color = color.strip().upper()
        if re.match(r'^#[0-9A-F]{6}$', color):
            formatted_color = color
        elif re.match(r'^#[0-9A-F]{3}$', color):
            formatted_color = f"#{color[1]}{color[1]}{color[2]}{color[2]}{color[3]}{color[3]}"
        elif re.match(r'^[0-9A-F]{6}$', color):
            formatted_color = f"#{color}"
        elif re.match(r'^[0-9A-F]{3}$', color):
            formatted_color = f"#{color[0]}{color[0]}{color[1]}{color[1]}{color[2]}{color[2]}"
        elif color in StyleConstants.COLOR_NAMES:
            formatted_color = StyleConstants.COLOR_NAMES[color]
        else:
            return StyleConstants.DEFAULT_FONT_COLOR if is_font_color else None

        if is_border_color:
            if formatted_color.upper() in ['#FFFFFF', '#FFF', 'WHITE', '#FEFEFE', '#FDFDFD']:
                return '#E0E0E0'
            elif formatted_color.upper() in ['#D8D8D8', '#DADADA', '#DBDBDB']:
                return '#E0E0E0'
        return formatted_color
