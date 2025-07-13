"""
XLSX解析器模块

解析Excel XLSX文件并转换为Sheet对象
包含完整的样式提取、颜色处理、边框识别等功能，支持流式读取。
"""

import openpyxl
from typing import Iterator, Optional
from src.models.table_model import Sheet, Row, Cell, Style, LazySheet, LazyRowProvider
from src.parsers.base_parser import BaseParser


class XlsxRowProvider:
    """Lazy row provider for XLSX files using openpyxl streaming with read_only=True."""
    
    def __init__(self, file_path: str, sheet_name: Optional[str] = None):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self._total_rows_cache: Optional[int] = None
        self._merged_cells_cache: Optional[list[str]] = None
        self._worksheet_title_cache: Optional[str] = None
    
    def _get_worksheet_info(self):
        """Get worksheet info without reading all data."""
        if self._worksheet_title_cache is None:
            workbook = openpyxl.load_workbook(self.file_path, read_only=True)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            self._worksheet_title_cache = worksheet.title
            workbook.close()
        return self._worksheet_title_cache
    
    def _get_merged_cells(self) -> list[str]:
        """Get merged cells info."""
        if self._merged_cells_cache is None:
            # Read merged cells from non-read-only workbook (required for merged_cells access)
            workbook = openpyxl.load_workbook(self.file_path)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            self._merged_cells_cache = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]
            workbook.close()
        return self._merged_cells_cache
    
    def _parse_row(self, row_cells: tuple) -> Row:
        """Parse a tuple of openpyxl cells into a Row object."""
        parser = XlsxParser()
        cells = []
        for cell in row_cells:
            cell_value = cell.value
            cell_style = parser._extract_style(cell) if cell else None
            cells.append(Cell(value=cell_value, style=cell_style))
        return Row(cells=cells)

    def iter_rows(self, start_row: int = 0, max_rows: Optional[int] = None) -> Iterator[Row]:
        """Yield rows on demand using openpyxl.iter_rows with read_only=True."""
        workbook = openpyxl.load_workbook(self.file_path, read_only=True)
        worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
        
        try:
            # Use openpyxl's iter_rows with values_only=False to get cell objects with styles
            row_count = 0
            for row_cells in worksheet.iter_rows(values_only=False):
                # Skip rows before start_row
                if row_count < start_row:
                    row_count += 1
                    continue
                
                # Stop if we've reached max_rows
                if max_rows is not None and (row_count - start_row) >= max_rows:
                    break
                
                yield self._parse_row(row_cells)
                row_count += 1
        
        finally:
            workbook.close()
    
    def get_row(self, row_index: int) -> Row:
        """Get a specific row by index."""
        workbook = openpyxl.load_workbook(self.file_path, read_only=True)
        worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
        
        try:
            # Get specific row using worksheet.iter_rows
            for i, row_cells in enumerate(worksheet.iter_rows(values_only=False)):
                if i == row_index:
                    return self._parse_row(row_cells)
            
            raise IndexError(f"Row index {row_index} out of range")
        
        finally:
            workbook.close()
    
    def get_total_rows(self) -> int:
        """Get total number of rows without loading all data."""
        if self._total_rows_cache is None:
            workbook = openpyxl.load_workbook(self.file_path, read_only=True)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            
            try:
                # Count rows using worksheet.max_row property
                self._total_rows_cache = worksheet.max_row or 0
            finally:
                workbook.close()
        
        return self._total_rows_cache


class XlsxParser(BaseParser):
    """
    A parser for XLSX files with comprehensive style extraction.

    This parser handles modern Excel files (.xlsx) and is capable of extracting
    a wide range of styling information, including fonts, colors, borders,
    and number formats. It also supports streaming for large files via the
    XlsxRowProvider.
    """

    def parse(self, file_path: str) -> Sheet:
        """
        Parses an XLSX file and returns a Sheet object.

        This method loads the entire file into memory to parse its content
        and styles. For large files, consider using `create_lazy_sheet`.

        Args:
            file_path: The absolute path to the XLSX file.

        Returns:
            A Sheet object containing the full data and styles.

        Raises:
            ValueError: If the workbook contains no active worksheet.
            FileNotFoundError: If the file does not exist.
        """
        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active

        if worksheet is None:
            raise ValueError("工作簿不包含任何活动工作表")

        rows = []
        for row in worksheet.iter_rows():
            cells = []
            for cell in row:
                # 提取单元格值和样式
                cell_value = cell.value
                cell_style = self._extract_style(cell)

                # 创建包含样式的Cell对象
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

    def _column_letter_to_number(self, column_letter: str) -> int:
        """
        将Excel列字母（如 'A', 'B', 'AA'）转换为数字（1-based）。
        """
        result = 0
        for char in column_letter:
            result = result * 26 + (ord(char.upper()) - ord('A') + 1)
        return result

    def _parse_range(self, range_string: str) -> tuple[int, int, int, int]:
        """
        解析Excel范围字符串（如 "A1:C3"）为数字索引。
        """
        import re
        match = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", range_string, re.IGNORECASE)
        if not match:
            raise ValueError(f"无效的范围格式: {range_string}")
        
        start_col_letter, start_row, end_col_letter, end_row = match.groups()
        
        start_col = self._column_letter_to_number(start_col_letter)
        end_col = self._column_letter_to_number(end_col_letter)
        
        return int(start_row), start_col, int(end_row), end_col

    def _extract_style(self, cell) -> Style:
        """
        Extracts comprehensive style information from an openpyxl cell.

        Args:
            cell: An openpyxl cell object.

        Returns:
            A Style object populated with the cell's formatting.
        """
        # 初始化默认样式
        style = Style()

        # 提取字体属性
        if cell.font:
            font = cell.font
            style.bold = font.bold if font.bold is not None else False
            style.italic = font.italic if font.italic is not None else False
            style.underline = font.underline is not None and font.underline != 'none'
            style.font_size = font.size if font.size else None
            style.font_name = font.name if font.name else None

            # 提取字体颜色（全面方法）
            if font.color:
                font_color = self._extract_color(font.color)
                if font_color and font_color != "#000000":  # 仅在非默认黑色时设置
                    style.font_color = font_color

        # 提取填充/背景属性（增强全面方法）
        if cell.fill:
            background_color = self._extract_fill_color(cell.fill)
            if background_color:
                style.background_color = background_color

        # 提取对齐属性
        if cell.alignment:
            alignment = cell.alignment
            # 水平对齐
            if alignment.horizontal:
                style.text_align = alignment.horizontal
            # 垂直对齐
            if alignment.vertical:
                style.vertical_align = alignment.vertical
            # 文本换行
            style.wrap_text = alignment.wrap_text if alignment.wrap_text is not None else False

        # 提取边框属性
        if cell.border:
            border = cell.border
            style.border_top = self._get_border_style(border.top) if border.top else ""
            style.border_bottom = self._get_border_style(border.bottom) if border.bottom else ""
            style.border_left = self._get_border_style(border.left) if border.left else ""
            style.border_right = self._get_border_style(border.right) if border.right else ""

            # 提取边框颜色（全面方法）
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
            else:
                style.border_color = Style().border_color

        # 提取数字格式（增强）
        style.number_format = self._extract_number_format(cell)

        # 提取超链接信息
        if cell.hyperlink:
            try:
                # 获取超链接目标
                if hasattr(cell.hyperlink, 'target'):
                    style.hyperlink = cell.hyperlink.target
                elif hasattr(cell.hyperlink, 'location'):
                    style.hyperlink = cell.hyperlink.location
            except:
                pass

        # 提取注释信息
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

    def _get_border_style(self, border_side) -> str:
        """
        将openpyxl边框样式转换为CSS边框样式。

        Args:
            border_side: openpyxl边框对象

        Returns:
            CSS边框样式字符串
        """
        if not border_side or not border_side.style:
            return ""

        # 将openpyxl边框样式映射到CSS等效样式
        border_style_map = {
            'thin': '1px solid',
            'medium': '2px solid',
            'thick': '2px solid',
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
        
        # 如果样式不在我们的映射中，也视为空
        style_str = border_style_map.get(border_side.style)
        if not style_str:
            return ""
        
        color = None
        if border_side.color:
            color = self._extract_color(border_side.color)

        # 如果没有提取到有效颜色，使用默认黑色
        final_color = color if color else "#000000"
        
        return f"{style_str} {final_color}"

    def _extract_fill_color(self, fill) -> str | None:
        """
        增强的填充颜色提取，支持多种填充类型。

        Args:
            fill: openpyxl填充对象

        Returns:
            提取的背景颜色，如果无有效颜色则返回None
        """
        if not fill:
            return None

        try:
            # 1. 实色填充 (PatternFill with solid pattern)
            if hasattr(fill, 'patternType') and fill.patternType:
                if fill.patternType == 'solid' and hasattr(fill, 'start_color') and fill.start_color:
                    color = self._extract_color(fill.start_color)
                    # 过滤掉默认的白色和黑色背景
                    if color and color not in ["#FFFFFF", "#000000"]:
                        return color

                # 2. 图案填充 (其他pattern类型)
                elif fill.patternType in ['lightGray', 'mediumGray', 'darkGray']:
                    pattern_colors = {
                        'lightGray': "#F2F2F2",
                        'mediumGray': "#D9D9D9",
                        'darkGray': "#BFBFBF"
                    }
                    return pattern_colors.get(fill.patternType)

                # 3. 其他图案填充，尝试提取前景色
                elif hasattr(fill, 'fgColor') and fill.fgColor:
                    color = self._extract_color(fill.fgColor)
                    if color and color not in ["#FFFFFF", "#000000"]:
                        return color

            # 4. 渐变填充 (GradientFill)
            if hasattr(fill, 'type') and fill.type == 'gradient':
                # 对于渐变，取第一个停止点的颜色作为主色
                if hasattr(fill, 'stop') and fill.stop:
                    # stop可能是列表或单个值
                    if isinstance(fill.stop, (list, tuple)) and len(fill.stop) > 0:
                        first_stop = fill.stop[0]
                        if hasattr(first_stop, 'color') and first_stop.color:
                            color = self._extract_color(first_stop.color)
                            if color and color not in ["#FFFFFF", "#000000"]:
                                return color
                    elif hasattr(fill.stop, 'color'):
                        color = self._extract_color(fill.stop.color)
                        if color and color not in ["#FFFFFF", "#000000"]:
                            return color

        except Exception:
            # 静默处理填充提取错误，避免中断解析流程
            pass

        return None

    def _extract_color(self, color_obj) -> str | None:
        """
        简化的颜色提取方法，使用openpyxl的统一value接口。

        Args:
            color_obj: openpyxl 颜色对象

        Returns:
            十六进制颜色字符串，如 "#FF0000"，失败时返回 None
        """
        if not color_obj:
            return None

        try:
            # 使用openpyxl的统一value接口
            value = color_obj.value

            if isinstance(value, str) and len(value) >= 6:
                # RGB/ARGB格式 - 取最后6位作为RGB
                return f"#{value[-6:]}"
            elif isinstance(value, int):
                # 索引或主题颜色 - 使用映射表
                return self._get_color_by_index(value)

        except Exception:
            # 静默处理错误，避免中断解析流程
            pass

        return None

    def _get_color_by_index(self, index: int) -> str:
        """
        统一的索引颜色映射，支持标准索引和主题颜色。

        Args:
            index: 颜色索引或主题索引

        Returns:
            十六进制颜色字符串
        """
        # 合并的颜色映射表（索引颜色 + 主题颜色）
        color_map = {
            # 标准索引颜色
            0: "#000000",   # 黑色
            1: "#FFFFFF",   # 白色
            2: "#FF0000",   # 红色
            3: "#00FF00",   # 绿色
            4: "#0000FF",   # 蓝色
            5: "#FFFF00",   # 黄色
            6: "#FF00FF",   # 洋红
            7: "#00FFFF",   # 青色
            8: "#000000",   # 黑色
            9: "#FFFFFF",   # 白色
            10: "#FF0000",  # 红色
            64: "#000000",  # 自动颜色（通常是黑色）

            # 主题颜色（索引0-9对应主题颜色）
            # 注意：这里可能与标准索引有重叠，以主题颜色为准
            # 0: "#FFFFFF",  # 背景1（与索引0重叠）
            # 1: "#000000",  # 文本1（与索引1重叠）
            # 2: "#E7E6E6",  # 背景2
            # 3: "#44546A",  # 文本2
            # 4: "#5B9BD5",  # 强调1
            # 5: "#70AD47",  # 强调2
            # 6: "#FFC000",  # 强调3
            # 7: "#264478",  # 强调4
            # 8: "#7030A0",  # 强调5
            # 9: "#0F243E",  # 强调6
        }

        return color_map.get(index, "#000000")  # 默认黑色

    def _extract_number_format(self, cell) -> str:
        """
        增强的数字格式提取。

        Args:
            cell: openpyxl单元格对象

        Returns:
            数字格式字符串
        """
        try:
            if cell.number_format and cell.number_format != 'General':
                # 标准化一些常见的数字格式
                format_str = cell.number_format

                # 处理一些特殊格式
                format_mappings = {
                    '0.00': '数字(2位小数)',
                    '0%': '百分比',
                    '0.00%': '百分比(2位小数)',
                    'mm/dd/yyyy': '日期(月/日/年)',
                    'dd/mm/yyyy': '日期(日/月/年)',
                    'yyyy-mm-dd': '日期(年-月-日)',
                    '$#,##0.00': '货币',
                    '¥#,##0.00': '人民币',
                    '#,##0': '千分位数字'
                }

                # 如果是常见格式，返回中文描述，否则返回原格式
                return format_mappings.get(format_str, format_str) or ""
        except:
            pass

        return ""

    def _extract_hyperlink(self, cell) -> str | None:
        """
        增强的超链接提取。

        Args:
            cell: openpyxl单元格对象

        Returns:
            超链接URL或None
        """
        if not cell.hyperlink:
            return None

        try:
            # 1. 外部链接 (target属性)
            if hasattr(cell.hyperlink, 'target') and cell.hyperlink.target:
                target = cell.hyperlink.target
                # 确保是有效的URL格式
                if isinstance(target, str) and (
                    target.startswith(('http://', 'https://', 'ftp://', 'mailto:')) or
                    target.startswith('file://') or
                    '.' in target  # 可能是文件路径或域名
                ):
                    return target

            # 2. 内部链接 (location属性)
            if hasattr(cell.hyperlink, 'location') and cell.hyperlink.location:
                location = cell.hyperlink.location
                if isinstance(location, str):
                    # 内部链接通常是工作表引用，添加前缀标识
                    return f"#内部:{location}"

            # 3. 其他类型的超链接
            if hasattr(cell.hyperlink, 'display') and cell.hyperlink.display:
                return str(cell.hyperlink.display)

        except Exception:
            pass

        return None
    
    def supports_streaming(self) -> bool:
        """XLSX parser supports streaming."""
        return True
    
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> LazySheet:
        """
        Creates a LazySheet for streaming data from an XLSX file.

        Args:
            file_path: The absolute path to the XLSX file.
            sheet_name: The name of the sheet to parse (optional).

        Returns:
            A LazySheet object that can stream data on demand.
        """
        provider = XlsxRowProvider(file_path, sheet_name)
        name = provider._get_worksheet_info()
        merged_cells = provider._get_merged_cells()
        return LazySheet(name=name, provider=provider, merged_cells=merged_cells)
