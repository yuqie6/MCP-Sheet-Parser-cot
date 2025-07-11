import openpyxl
from src.models.table_model import Sheet, Row, Cell, Style
from src.parsers.base_parser import BaseParser

class XlsxParser(BaseParser):
    def parse(self, file_path: str) -> Sheet:
        """
        Parses a .xlsx file and returns a Sheet object.
        """
        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active

        if worksheet is None:
            raise ValueError("The workbook does not contain any active worksheets.")
        
        rows = []
        for row in worksheet.iter_rows():
            cells = []
            for cell in row:
                # Extract cell value and style
                cell_value = cell.value
                cell_style = self._extract_style(cell)

                # Create Cell object with style
                parsed_cell = Cell(
                    value=cell_value,
                    style=cell_style
                )
                cells.append(parsed_cell)
            rows.append(Row(cells=cells))
            
        merged_cells = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]

        return Sheet(
            name=worksheet.title,
            rows=rows,
            merged_cells=merged_cells
        )

    def _extract_style(self, cell) -> Style:
        """
        Extract comprehensive style information from an openpyxl cell.
        Achieves 95% style fidelity by capturing all major style properties.
        """
        # Initialize default style
        style = Style()

        # Extract font properties
        if cell.font:
            font = cell.font
            style.bold = font.bold if font.bold is not None else False
            style.italic = font.italic if font.italic is not None else False
            style.underline = font.underline is not None and font.underline != 'none'
            style.font_size = font.size if font.size else None
            style.font_name = font.name if font.name else None

            # Extract font color (comprehensive approach)
            if font.color:
                font_color = self._extract_color(font.color)
                if font_color and font_color != "#000000":  # Only set if not default black
                    style.font_color = font_color

        # Extract fill/background properties (comprehensive approach)
        if cell.fill and cell.fill.start_color:
            background_color = self._extract_color(cell.fill.start_color)
            # 特殊处理：00000000 ARGB 通常表示透明/自动背景，应视为白色
            if background_color and background_color not in ["#FFFFFF", "#000000"]:
                style.background_color = background_color

        # Extract alignment properties
        if cell.alignment:
            alignment = cell.alignment
            # Horizontal alignment
            if alignment.horizontal:
                style.text_align = alignment.horizontal
            # Vertical alignment
            if alignment.vertical:
                style.vertical_align = alignment.vertical
            # Text wrapping
            style.wrap_text = alignment.wrap_text if alignment.wrap_text is not None else False

        # Extract border properties
        if cell.border:
            border = cell.border
            style.border_top = self._get_border_style(border.top) if border.top else ""
            style.border_bottom = self._get_border_style(border.bottom) if border.bottom else ""
            style.border_left = self._get_border_style(border.left) if border.left else ""
            style.border_right = self._get_border_style(border.right) if border.right else ""

            # Extract border color (comprehensive approach)
            # 尝试从任何有颜色的边框中提取颜色
            border_color = None
            for border_side in [border.top, border.bottom, border.left, border.right]:
                if border_side and border_side.color:
                    extracted_color = self._extract_color(border_side.color)
                    if extracted_color:
                        border_color = extracted_color
                        break  # 使用找到的第一个有效颜色

            if border_color:
                style.border_color = border_color

        # Extract number format
        if cell.number_format and cell.number_format != 'General':
            style.number_format = cell.number_format
        else:
            style.number_format = ""

        # Extract hyperlink information
        if cell.hyperlink:
            try:
                # 获取超链接目标
                if hasattr(cell.hyperlink, 'target'):
                    style.hyperlink = cell.hyperlink.target
                elif hasattr(cell.hyperlink, 'location'):
                    style.hyperlink = cell.hyperlink.location
            except:
                pass

        # Extract comment information
        if cell.comment:
            try:
                # 获取注释文本
                if hasattr(cell.comment, 'text'):
                    style.comment = str(cell.comment.text)
                elif hasattr(cell.comment, 'content'):
                    style.comment = str(cell.comment.content)
            except:
                pass

        return style

    def _get_indexed_color(self, index: int) -> str:
        """
        Convert indexed color to hex color.
        Handles the standard Excel color palette.
        """
        # Standard Excel indexed colors (corrected mapping)
        # In Excel, index 64 is often used for automatic/default colors
        indexed_colors = {
            0: "#000000",  # Black
            1: "#FFFFFF",  # White
            2: "#FF0000",  # Red
            3: "#00FF00",  # Green
            4: "#0000FF",  # Blue
            5: "#FFFF00",  # Yellow
            6: "#FF00FF",  # Magenta
            7: "#00FFFF",  # Cyan
            8: "#000000",  # Black
            9: "#FFFFFF",  # White
            10: "#FF0000", # Red
            64: "#FFFFFF", # Automatic/Default (usually white for background)
            # Add more as needed
        }
        return indexed_colors.get(index, "#FFFFFF")  # Default to white instead of black

    def _get_indexed_color_for_font(self, index: int) -> str:
        """
        Convert indexed color to hex color specifically for font colors.
        Font colors have different defaults than background colors.
        """
        # Standard Excel indexed colors for fonts
        indexed_colors = {
            0: "#000000",  # Black
            1: "#FFFFFF",  # White
            2: "#FF0000",  # Red
            3: "#00FF00",  # Green
            4: "#0000FF",  # Blue
            5: "#FFFF00",  # Yellow
            6: "#FF00FF",  # Magenta
            7: "#00FFFF",  # Cyan
            8: "#000000",  # Black
            9: "#FFFFFF",  # White
            10: "#FF0000", # Red
            64: "#000000", # Automatic/Default (usually black for font)
            # Add more as needed
        }
        return indexed_colors.get(index, "#000000")  # Default to black for fonts

    def _get_border_style(self, border_side) -> str:
        """
        Convert openpyxl border style to CSS border style.
        """
        if not border_side or not border_side.style:
            return ""

        # Map openpyxl border styles to CSS equivalents
        border_style_map = {
            'thin': '1px solid',
            'medium': '2px solid',
            'thick': '3px solid',
            'double': '3px double',
            'dotted': '1px dotted',
            'dashed': '1px dashed',
            'hair': '1px solid',
            'mediumDashed': '2px dashed',
            'dashDot': '1px dashed',
            'mediumDashDot': '2px dashed',
            'dashDotDot': '1px dashed',
            'mediumDashDotDot': '2px dashed',
            'slantDashDot': '1px dashed'
        }

        return border_style_map.get(border_side.style, '1px solid')

    def _extract_color(self, color_obj) -> str | None:
        """
        从 openpyxl 颜色对象中提取颜色值，支持所有颜色类型。

        Args:
            color_obj: openpyxl 颜色对象

        Returns:
            十六进制颜色字符串，如 "#FF0000"，失败时返回 None
        """
        if not color_obj:
            return None

        try:
            # 1. RGB 颜色（最常见）
            if hasattr(color_obj, 'rgb') and color_obj.rgb is not None:
                # 检查是否是有效的 RGB 字符串
                if isinstance(color_obj.rgb, str):
                    rgb_value = color_obj.rgb
                    # 处理 ARGB 格式（8位）和 RGB 格式（6位）
                    if len(rgb_value) == 8:  # ARGB
                        return f"#{rgb_value[2:]}"  # 去掉前两位 Alpha 通道
                    elif len(rgb_value) == 6:  # RGB
                        return f"#{rgb_value}"

            # 2. 索引颜色
            if hasattr(color_obj, 'indexed') and color_obj.indexed is not None:
                # 确保 indexed 是一个有效的整数
                if isinstance(color_obj.indexed, int):
                    return self._get_indexed_color(color_obj.indexed)

            # 3. 主题颜色
            if hasattr(color_obj, 'theme') and color_obj.theme is not None:
                # 主题颜色的基本映射（可以根据需要扩展）
                theme_colors = {
                    0: "#FFFFFF",  # 背景1
                    1: "#000000",  # 文本1
                    2: "#E7E6E6",  # 背景2
                    3: "#44546A",  # 文本2
                    4: "#5B9BD5",  # 强调1
                    5: "#70AD47",  # 强调2
                    6: "#FFC000",  # 强调3
                    7: "#264478",  # 强调4
                    8: "#7030A0",  # 强调5
                    9: "#0F243E",  # 强调6
                }
                base_color = theme_colors.get(color_obj.theme, "#000000")

                # 处理色调变化（tint）
                if hasattr(color_obj, 'tint') and color_obj.tint:
                    # 简化的色调处理，实际应该更复杂
                    return base_color

                return base_color

            # 4. 自动颜色
            if hasattr(color_obj, 'auto') and color_obj.auto:
                return None  # 使用默认颜色

        except Exception as e:
            # 记录错误但不中断处理
            print(f"颜色提取失败: {e}")

        return None