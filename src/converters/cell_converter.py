from datetime import datetime as dt, timedelta

from src.models.table_model import Cell, RichTextFragment

# Number format mapping constant
NUMBER_FORMAT_MAP = {
    "General": lambda v: str(v),
    "0": lambda v: f"{v:.0f}" if isinstance(v, (int, float)) else str(v),
    "0.0": lambda v: f"{v:.1f}" if isinstance(v, (int, float)) else str(v),
    "0.00": lambda v: f"{v:.2f}" if isinstance(v, (int, float)) else str(v),
    "#,##0": lambda v: f"{v:,.0f}" if isinstance(v, (int, float)) else str(v),
    "#,##0.0": lambda v: f"{v:,.1f}" if isinstance(v, (int, float)) else str(v),
    "#,##0.00": lambda v: f"{v:,.2f}" if isinstance(v, (int, float)) else str(v),
    "0%": lambda v: f"{v:.0%}" if isinstance(v, (int, float)) else str(v),
    "0.0%": lambda v: f"{v:.1%}" if isinstance(v, (int, float)) else str(v),
    "0.00%": lambda v: f"{v:.2%}" if isinstance(v, (int, float)) else str(v),
}

# Date format mapping constant
DATE_FORMAT_MAP = {
    "yyyy-mm-dd": "%Y-%m-%d",
    "mm/dd/yyyy": "%m/%d/%Y",
    "dd/mm/yyyy": "%d/%m/%Y",
    "yyyy/mm/dd": "%Y/%m/%d",
    "mm-dd-yyyy": "%m-%d-%Y",
    "dd-mm-yyyy": "%d-%m-%Y",
}


def format_chinese_date(date_obj: dt, format_str: str) -> str:
    """Formats a Chinese date."""
    if 'm"月"d"日"' in format_str:
        return f"{date_obj.month}月{date_obj.day}日"
    elif 'yyyy"年"m"月"d"日"' in format_str:
        return f"{date_obj.year}年{date_obj.month}月{date_obj.day}日"
    else:
        return f"{date_obj.month}月{date_obj.day}日"


class CellConverter:
    """Handles the HTML content generation for a single cell."""

    def __init__(self, style_converter):
        self.style_converter = style_converter

    def convert(self, cell: Cell) -> str:
        """
        Converts a Cell object to its HTML representation.
        """
        if isinstance(cell.value, list):
            return self._format_rich_text(cell.value)
        if cell.value is None:
            return ""
        if cell.style and cell.style.number_format:
            try:
                return self._apply_number_format(cell.value, cell.style.number_format)
            except Exception:
                pass
        if isinstance(cell.value, float):
            return f"{cell.value:.2f}".rstrip('0').rstrip('.') if not cell.value.is_integer() else str(int(cell.value))
        return str(cell.value)

    def _format_rich_text(self, fragments: list[RichTextFragment]) -> str:
        """
        Formats rich text fragments into a single HTML string.
        """
        return "".join(self._format_rich_text_fragment(f) for f in fragments)

    def _format_rich_text_fragment(self, fragment: RichTextFragment) -> str:
        """
        Formats a single rich text fragment into a styled HTML span.
        """
        style = fragment.style
        css_parts = []
        if style.font_name:
            css_parts.append(f"font-family: {self.style_converter._format_font_family(style.font_name)};")
        if style.font_size:
            css_parts.append(f"font-size: {self.style_converter._format_font_size(style.font_size)};")
        if style.font_color:
            css_parts.append(f"color: {self.style_converter._format_color(style.font_color, is_font_color=True)};")
        if style.bold:
            css_parts.append("font-weight: bold;")
        if style.italic:
            css_parts.append("font-style: italic;")
        if style.underline:
            css_parts.append("text-decoration: underline;")

        style_attr = f'style="{" ".join(css_parts)}"' if css_parts else ""
        return f'<span {style_attr}>{self._escape_html(fragment.text)}</span>'

    def _apply_number_format(self, value, number_format: str) -> str:
        """
        Applies a number format to a value.
        """
        if number_format in NUMBER_FORMAT_MAP:
            return NUMBER_FORMAT_MAP[number_format](value)
        if isinstance(value, (int, float)) and ("月" in number_format and "日" in number_format):
            try:
                excel_epoch = dt(1899, 12, 30)
                date_obj = excel_epoch + timedelta(days=value)
                return format_chinese_date(date_obj, number_format)
            except Exception:
                pass
        if isinstance(value, dt):
            if "月" in number_format and "日" in number_format:
                return format_chinese_date(value, number_format)
            if "yyyy" in number_format.lower() or "mm" in number_format.lower() or "dd" in number_format.lower():
                format_lower = number_format.lower()
                for excel_fmt, formatter in DATE_FORMAT_MAP.items():
                    if excel_fmt in format_lower:
                        return value.strftime(str(formatter))
                return value.strftime("%Y-%m-%d")
        if isinstance(value, (int, float)) and "%" in number_format:
            return f"{value * 100:.1f}%"
        if isinstance(value, (int, float)) and "," in number_format:
            return f"{value:,.2f}"
        return str(value)

    def _escape_html(self, text: str) -> str:
        """
        Escapes special HTML characters.
        """
        if not isinstance(text, str):
            text = str(text)
        return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&#x27;'))
