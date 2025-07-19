import logging

from src.constants import StyleConstants
from src.font_manager import get_font_manager
from src.models.table_model import Sheet, Style
from src.utils.color_utils import format_color

logger = logging.getLogger(__name__)


class StyleConverter:
    """处理所有 CSS 生成逻辑。"""

    def collect_styles(self, sheet: Sheet) -> dict[str, Style]:
        """
        收集表格中的所有唯一样式。
        参数：
            sheet: Sheet对象。
        返回：
            以样式ID为键、Style对象为值的字典。
        """
        styles = {}
        seen_style_keys = set()
        style_counter = 0

        for row in sheet.rows:
            for cell in row.cells:
                if cell.style:
                    style_key = self.get_style_key(cell.style)
                    if style_key not in seen_style_keys:
                        seen_style_keys.add(style_key)
                        style_id = f"style_{style_counter}"
                        styles[style_id] = cell.style
                        style_counter += 1
        return styles

    def get_style_key(self, style: Style) -> str:
        """
        生成样式的唯一标识符。
        参数：
            style: Style对象。
        返回：
            唯一字符串标识。
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

    def generate_css(self, styles: dict[str, Style], sheet: Sheet | None = None) -> str:
        """
        生成 CSS 样式。
        参数：
            styles: 样式字典。
            sheet: Sheet对象，用于维度信息。
        返回：CSS字符串。
        """
        css_rules = [
            """
        body {
            margin: 20px;
            font-family: Arial, sans-serif;
        }
        h1 {
            font-size: 24px;
            margin: 20px 0;
            color: #333;
            text-align: center;
        }
        table {
            border-collapse: separate;
            border-spacing: 0;
            font-family: Arial, sans-serif;
            table-layout: auto;
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

        /* 为溢出单元格提供例外 */
        td.text-overflow {
            word-wrap: normal !important;
            white-space: nowrap !important;
            min-width: auto !important;
            max-width: none !important;
            width: auto !important;
        }

        /* 覆盖列宽限制，确保溢出单元格能够扩展 */
        table td.text-overflow {
            width: auto !important;
            min-width: auto !important;
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
            position: relative;
            overflow: visible;
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
                formatted_color = format_color(style.font_color, is_font_color=True)
                css_rule += f" color: {formatted_color} !important;"
            if style.bold:
                css_rule += " font-weight: bold;"
            if style.italic:
                css_rule += " font-style: italic;"
            if style.underline:
                css_rule += " text-decoration: underline;"
            if style.background_color:
                formatted_bg_color = format_color(style.background_color, is_font_color=False,
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
        css_rules.append(self._generate_text_overflow_css())

        return "\n".join(css_rules)

    def _generate_text_overflow_css(self) -> str:
        """
        生成文字溢出显示的CSS样式（模拟Excel行为）。
        """
        return """
        /* Excel文字溢出显示效果 - 使用最高优先级 */
        table tr td.text-overflow,
        table tbody tr td.text-overflow,
        td.text-overflow {
            overflow: visible !important;
            white-space: nowrap !important;
            position: relative !important;
            z-index: 5 !important;
            max-width: none !important;
            width: auto !important;
            word-wrap: normal !important;
            min-width: auto !important;
        }

        /* 覆盖所有可能的列宽限制 */
        table tr td.text-overflow,
        table tbody tr td.text-overflow,
        table td.text-overflow {
            width: auto !important;
            min-width: auto !important;
        }

        /* 覆盖nth-child列宽限制 - 扩展到所有列 */
        table td:nth-child(n).text-overflow {
            width: auto !important;
            min-width: auto !important;
            max-width: none !important;
            white-space: nowrap !important;
            word-wrap: normal !important;
            overflow: visible !important;
            position: relative !important;
            z-index: 5 !important;
        }

        /* 确保溢出文字显示在其他单元格之上 */
        td.text-overflow:hover {
            z-index: 10 !important;
        }

        /* 确保表格单元格不会裁剪溢出内容 */
        table {
            table-layout: auto !important;
        }
        """

    def _generate_dimension_css(self, sheet: Sheet) -> str:
        """
        生成列宽和行高的CSS。
        """
        css_rules = []
        excel_to_px = 8.43
        for col_idx, width in sheet.column_widths.items():
            if width > 0:
                width_px = width * excel_to_px
                css_rules.append(
                    f"table td:nth-child({col_idx + 1}):not(.text-overflow), table th:nth-child({col_idx + 1}) {{ width: {width_px:.0f}px; min-width: {width_px:.0f}px; }}")
        for row_idx, height in sheet.row_heights.items():
            if height > 0:
                # 确保行高至少能容纳内容，避免文字被截断
                min_height = max(height, 18.0)  # 最小18pt行高
                css_rules.append(f"table tr:nth-child({row_idx + 1}) {{ height: {min_height}pt; min-height: {min_height}pt; }}")
        return "\n".join(css_rules)

    def _generate_chart_css(self) -> str:
        """生成图表相关CSS。"""
        return """
        .chart-overlay {
            position: absolute;
            z-index: 10;
            pointer-events: auto;
            background: none;
            border: none;
            border-radius: 0;
            box-shadow: none;
            padding: 0;
        }
        .chart-overlay:hover {
            z-index: 20;
            box-shadow: none;
        }
        .chart-container {
            margin: 20px 0;
            padding: 0;
            border: none;
            border-radius: 0;
            background-color: transparent;
            box-shadow: none;
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
        生成边框相关CSS。
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



    def _format_font_family(self, font_name: str) -> str:
        """
        格式化字体族名称。
        """
        font_manager = get_font_manager()
        return font_manager.generate_font_family(font_name)

    def _format_font_size(self, font_size: float) -> str:
        """
        格式化字体大小。
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
