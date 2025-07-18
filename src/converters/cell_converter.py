from datetime import datetime as dt, timedelta

from src.models.table_model import Cell, RichTextFragment
from src.utils.html_utils import escape_html, create_html_element
from src.utils.color_utils import format_color

 # 数字格式映射常量
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

 # 日期格式映射常量
DATE_FORMAT_MAP = {
    "yyyy-mm-dd": "%Y-%m-%d",
    "mm/dd/yyyy": "%m/%d/%Y",
    "dd/mm/yyyy": "%d/%m/%Y",
    "yyyy/mm/dd": "%Y/%m/%d",
    "mm-dd-yyyy": "%m-%d-%Y",
    "dd-mm-yyyy": "%d-%m-%Y",
}


def format_chinese_date(date_obj: dt, format_str: str) -> str:
    """格式化中文日期。"""
    if 'm"月"d"日"' in format_str:
        return f"{date_obj.month}月{date_obj.day}日"
    elif 'yyyy"年"m"月"d"日"' in format_str:
        return f"{date_obj.year}年{date_obj.month}月{date_obj.day}日"
    else:
        return f"{date_obj.month}月{date_obj.day}日"


class CellConverter:
    """处理单元格的 HTML 内容生成。"""

    def __init__(self, style_converter):
        self.style_converter = style_converter

    def convert(self, cell: Cell) -> str:
        """
        将 Cell 对象转换为 HTML 表现形式。
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
        将富文本片段格式化为单一 HTML 字符串。
        """
        return "".join(self._format_rich_text_fragment(f) for f in fragments)

    def _format_rich_text_fragment(self, fragment: RichTextFragment) -> str:
        """
        将单个富文本片段格式化为带样式的 HTML span。
        """
        style = fragment.style
        inline_styles = {}
        
        if style.font_name:
            inline_styles['font-family'] = self.style_converter._format_font_family(style.font_name)
        if style.font_size:
            inline_styles['font-size'] = self.style_converter._format_font_size(style.font_size)
        if style.font_color:
            inline_styles['color'] = format_color(style.font_color, is_font_color=True)
        if style.bold:
            inline_styles['font-weight'] = 'bold'
        if style.italic:
            inline_styles['font-style'] = 'italic'
        if style.underline:
            inline_styles['text-decoration'] = 'underline'

        return create_html_element('span', escape_html(fragment.text), inline_styles=inline_styles)

    def _apply_number_format(self, value, number_format: str) -> str:
        """
        对值应用数字格式。
        """
        if number_format in NUMBER_FORMAT_MAP:
            return NUMBER_FORMAT_MAP[number_format](value)
        if isinstance(value, (int, float)) and ("月" in number_format and "日" in number_format):
            try:
                # 使用更精确的Excel日期转换，避免精度损失
                from decimal import Decimal
                excel_epoch = dt(1899, 12, 30)
                # 使用Decimal保持精度，然后转换为timedelta
                days_decimal = Decimal(str(value))
                days_int = int(days_decimal)
                microseconds = int((days_decimal - days_int) * 86400 * 1000000)
                date_obj = excel_epoch + timedelta(days=days_int, microseconds=microseconds)
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
