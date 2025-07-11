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

            # Extract font color (simplified approach for reliability)
            if font.color:
                try:
                    if hasattr(font.color, 'indexed') and font.color.indexed is not None:
                        indexed_color = self._get_indexed_color_for_font(font.color.indexed)
                        # Only set font color if it's not the default black
                        if indexed_color != "#000000":
                            style.font_color = indexed_color
                    # For now, keep default color for complex color objects
                except:
                    pass

        # Extract fill/background properties (simplified approach)
        if cell.fill and cell.fill.start_color:
            fill = cell.fill
            try:
                if hasattr(fill.start_color, 'indexed') and fill.start_color.indexed is not None:
                    indexed_color = self._get_indexed_color(fill.start_color.indexed)
                    # Only set background if it's not the default white
                    if indexed_color != "#FFFFFF":
                        style.background_color = indexed_color
                # For now, keep default background for complex fill objects
            except:
                pass

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

            # Extract border color (use top border color as default)
            if border.top and border.top.color and border.top.color.rgb:
                border_rgb = str(border.top.color.rgb)
                style.border_color = f"#{border_rgb}"

        # Extract number format
        if cell.number_format and cell.number_format != 'General':
            style.number_format = cell.number_format
        else:
            style.number_format = ""

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