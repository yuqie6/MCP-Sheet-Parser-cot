"""
XLS格式解析器模块

基于xlrd库实现XLS格式的完整解析，支持样式提取和数据转换。
"""

import logging
import xlrd
import xlrd.xldate
from src.models.table_model import Sheet, Row, Cell, Style, LazySheet
from src.parsers.base_parser import BaseParser
from src.utils.border_utils import get_xls_border_style_name

logger = logging.getLogger(__name__)


class XlsParser(BaseParser):
    """XLS格式解析器，基于xlrd库实现完整的样式提取。"""
    
    def __init__(self):
        # 默认Excel调色板（作为回退方案）
        self.default_color_map = {
            0: "#000000",   # 黑色
            1: "#FFFFFF",   # 白色
            2: "#FF0000",   # 红色
            3: "#00FF00",   # 绿色
            4: "#0000FF",   # 蓝色
            5: "#FFFF00",   # 黄色
            6: "#FF00FF",   # 洋红
            7: "#00FFFF",   # 青色
            8: "#800000",   # 深红
            9: "#008000",   # 深绿
            10: "#000080",  # 深蓝
            11: "#808000",  # 橄榄色
            12: "#800080",  # 紫色
            13: "#008080",  # 蓝绿色
            14: "#C0C0C0",  # 银色
            15: "#808080",  # 灰色
            16: "#9999FF",  # 浅蓝
            17: "#993366",  # 深粉
            18: "#FFFFCC",  # 浅黄
            19: "#CCFFFF",  # 浅青
            20: "#660066",  # 深紫
            21: "#FF8080",  # 浅红
            22: "#0066CC",  # 蓝色
            23: "#CCCCFF",  # 很浅蓝
        }
        # 动态颜色缓存
        self.workbook_colors = {}
    
    def parse(self, file_path: str) -> list[Sheet]:
        """
        解析XLS文件并返回Sheet对象列表。

        Args:
            file_path: XLS文件路径

        Returns:
            包含完整数据和样式的Sheet对象列表

        Raises:
            RuntimeError: 当解析失败时
        """
        try:
            # 打开XLS文件，启用格式化信息
            workbook = xlrd.open_workbook(file_path, formatting_info=True)

            # 检查工作表数量
            if workbook.nsheets == 0:
                raise RuntimeError("工作簿不包含任何工作表")

            sheets = []

            # 解析所有工作表
            for sheet_idx in range(workbook.nsheets):
                worksheet = workbook.sheet_by_index(sheet_idx)
                sheet_name = worksheet.name

                # 解析所有行和单元格
                rows = []
                for row_idx in range(worksheet.nrows):
                    cells = []
                    for col_idx in range(worksheet.ncols):
                        # 获取单元格值
                        cell_value = self._get_cell_value(workbook, worksheet, row_idx, col_idx)

                        # 提取样式信息
                        cell_style = self._extract_style(workbook, worksheet, row_idx, col_idx)

                        # 创建Cell对象
                        cell = Cell(
                            value=cell_value,
                            style=cell_style
                        )
                        cells.append(cell)

                    rows.append(Row(cells=cells))

                # 处理合并单元格
                merged_cells = self._extract_merged_cells(worksheet)

                sheet = Sheet(
                    name=sheet_name,
                    rows=rows,
                    merged_cells=merged_cells
                )
                sheets.append(sheet)

            return sheets

        except Exception as e:
            logger.error(f"解析XLS文件失败: {e}")
            raise RuntimeError(f"无法解析XLS文件 {file_path}: {str(e)}")
    
    def _get_cell_value(self, workbook, worksheet, row_idx: int, col_idx: int):
        """获取单元格的值，处理不同的数据类型。"""
        try:
            cell = worksheet.cell(row_idx, col_idx)
            cell_type = cell.ctype
            cell_value = cell.value
            
            # 处理不同的单元格类型
            if cell_type == xlrd.XL_CELL_EMPTY:
                return None
            elif cell_type == xlrd.XL_CELL_TEXT:
                return str(cell_value)
            elif cell_type == xlrd.XL_CELL_NUMBER:
                # xlrd 2.0版本中，需要手动检查是否为日期
                # 通过检查单元格的数字格式来判断是否为日期
                try:
                    # 获取单元格的格式信息
                    xf_index = cell.xf_index
                    if xf_index < len(workbook.xf_list):
                        xf = workbook.xf_list[xf_index]
                        if xf.format_key in workbook.format_map:
                            format_info = workbook.format_map[xf.format_key]
                            if format_info and format_info.format_str:
                                # 简单的日期格式检测
                                format_str = format_info.format_str.lower()
                                if any(date_indicator in format_str for date_indicator in
                                      ['d', 'm', 'y', 'h', 's', '/']):
                                    logger.debug(f"将单元格 ({row_idx}, {col_idx}) 的值 '{cell_value}'（格式：'{format_str}'）作为日期处理。")
                                    return xlrd.xldate.xldate_as_datetime(cell_value, worksheet.book.datemode)
                except (ValueError, TypeError, IndexError):
                    # 如果日期检测失败，当作普通数字处理
                    pass

                # 如果是整数，返回int，否则返回float
                return int(cell_value) if cell_value.is_integer() else cell_value
            elif cell_type == xlrd.XL_CELL_DATE:
                return xlrd.xldate.xldate_as_datetime(cell_value, worksheet.book.datemode)
            elif cell_type == xlrd.XL_CELL_BOOLEAN:
                return bool(cell_value)
            elif cell_type == xlrd.XL_CELL_ERROR:
                return f"#ERROR:{cell_value}"
            else:
                return cell_value
                
        except Exception as e:
            logger.warning(f"获取单元格值失败 ({row_idx}, {col_idx}): {e}")
            return None
    
    def _extract_style(self, workbook, worksheet, row_idx: int, col_idx: int) -> Style:
        """
        从XLS单元格提取样式信息。
        
        Args:
            workbook: xlrd工作簿对象
            worksheet: xlrd工作表对象
            row_idx: 行索引
            col_idx: 列索引
            
        Returns:
            Style对象
        """
        style = Style()
        
        try:
            # 获取单元格的格式索引
            cell = worksheet.cell(row_idx, col_idx)
            xf_index = cell.xf_index
            
            if xf_index >= len(workbook.xf_list):
                return style
                
            # 获取扩展格式记录
            xf = workbook.xf_list[xf_index]
            
            # 提取字体信息
            if xf.font_index < len(workbook.font_list):
                font = workbook.font_list[xf.font_index]
                style.bold = bool(font.bold)
                style.italic = bool(font.italic)
                style.underline = bool(font.underline_type)
                
                # 字体大小（xlrd中以20分之一点为单位）
                if font.height:
                    style.font_size = font.height / 20.0
                
                # 字体名称
                if font.name:
                    style.font_name = font.name
                
                # 字体颜色
                if font.colour_index:
                    style.font_color = self._get_color_from_index(workbook, font.colour_index)
            
            # 提取背景颜色（增强版）
            if hasattr(xf, 'background') and xf.background:
                # 获取填充模式
                fill_pattern = getattr(xf.background, 'fill_pattern', 0)

                if hasattr(xf.background, 'pattern_colour_index'):
                    bg_color_index = xf.background.pattern_colour_index
                    # 64是默认/自动颜色，通常是白色，我们将其视为空
                    if bg_color_index != 64:
                        bg_color = self._get_color_from_index(workbook, bg_color_index)
                        if bg_color:
                            style.background_color = bg_color

                # 如果有背景颜色索引，也尝试提取
                if hasattr(xf.background, 'background_colour_index'):
                    bg_color_index = xf.background.background_colour_index
                    if bg_color_index != 64 and not style.background_color:
                        bg_color = self._get_color_from_index(workbook, bg_color_index)
                        if bg_color:
                            style.background_color = bg_color
            
            # 提取对齐方式
            if hasattr(xf, 'alignment'):
                alignment = xf.alignment
                
                # 水平对齐
                if alignment.hor_align == 1:
                    style.text_align = "left"
                elif alignment.hor_align == 2:
                    style.text_align = "center"
                elif alignment.hor_align == 3:
                    style.text_align = "right"
                elif alignment.hor_align == 4:
                    style.text_align = "justify"
                
                # 垂直对齐
                if alignment.vert_align == 0:
                    style.vertical_align = "top"
                elif alignment.vert_align == 1:
                    style.vertical_align = "middle"
                elif alignment.vert_align == 2:
                    style.vertical_align = "bottom"
                
                # 文本换行
                style.wrap_text = bool(alignment.wrap)
            
            # 提取数字格式
            if xf.format_key < len(workbook.format_map):
                format_info = workbook.format_map[xf.format_key]
                if format_info and format_info.format_str:
                    style.number_format = format_info.format_str
            
            # 提取边框信息（增强版）
            if hasattr(xf, 'border'):
                border = xf.border
                if border:
                    # 处理各个边框 - 使用统一的边框工具
                    if border.top_line_style:
                        style_name = get_xls_border_style_name(border.top_line_style)
                        style.border_top = style_name if style_name else "solid"
                    if border.bottom_line_style:
                        style_name = get_xls_border_style_name(border.bottom_line_style)
                        style.border_bottom = style_name if style_name else "solid"
                    if border.left_line_style:
                        style_name = get_xls_border_style_name(border.left_line_style)
                        style.border_left = style_name if style_name else "solid"
                    if border.right_line_style:
                        style_name = get_xls_border_style_name(border.right_line_style)
                        style.border_right = style_name if style_name else "solid"

                    # 边框颜色（使用第一个有效的边框颜色）
                    for color_idx in [border.top_colour_index, border.bottom_colour_index,
                                    border.left_colour_index, border.right_colour_index]:
                        if color_idx and color_idx != 64: # 64是默认颜色
                            border_color = self._get_color_from_index(workbook, color_idx)
                            if border_color:
                                style.border_color = border_color
                                break
            
        except Exception as e:
            logger.warning(f"提取样式失败 ({row_idx}, {col_idx}): {e}")
        
        return style
    
    def _get_color_from_index(self, workbook, color_index: int) -> str:
        """
        将XLS颜色索引转换为RGB十六进制字符串。
        优先从工作簿的实际调色板获取，回退到默认调色板。

        Args:
            workbook: xlrd工作簿对象
            color_index: XLS颜色索引

        Returns:
            RGB颜色字符串，如 "#FF0000"
        """
        # 首先尝试从工作簿的调色板获取
        try:
            if hasattr(workbook, 'colour_map') and color_index in workbook.colour_map:
                rgb_tuple = workbook.colour_map[color_index]
                if rgb_tuple and len(rgb_tuple) >= 3:
                    r, g, b = rgb_tuple[:3]
                    return f"#{r:02X}{g:02X}{b:02X}"
        except Exception as e:
            logger.debug(f"从工作簿调色板获取颜色失败: {e}")

        # 回退到默认调色板
        if color_index in self.default_color_map:
            return self.default_color_map[color_index]

        # 对于完全未知的颜色索引，返回黑色
        logger.debug(f"未知颜色索引: {color_index}")
        return "#000000"
    
    def _extract_merged_cells(self, worksheet) -> list[str]:
        """
        提取合并单元格信息。
        
        Args:
            worksheet: xlrd工作表对象
            
        Returns:
            合并单元格范围列表，格式如 ["A1:B2", "C3:D4"]
        """
        merged_cells = []
        
        try:
            for crange in worksheet.merged_cells:
                # crange格式: (row_start, row_end, col_start, col_end)
                row_start, row_end, col_start, col_end = crange
                
                # 转换为Excel格式的范围字符串
                start_cell = self._index_to_excel_cell(row_start, col_start)
                end_cell = self._index_to_excel_cell(row_end - 1, col_end - 1)
                
                merged_cells.append(f"{start_cell}:{end_cell}")
                
        except Exception as e:
            logger.warning(f"提取合并单元格失败: {e}")
        
        return merged_cells
    
    def _index_to_excel_cell(self, row: int, col: int) -> str:
        """
        将行列索引转换为Excel单元格引用格式。
        
        Args:
            row: 行索引（0开始）
            col: 列索引（0开始）
            
        Returns:
            Excel单元格引用，如 "A1", "B2"
        """
        # 转换列索引为字母
        col_str = ""
        col += 1  # Excel列从1开始
        while col > 0:
            col -= 1
            col_str = chr(65 + col % 26) + col_str
            col //= 26
        
        # 行号从1开始
        return f"{col_str}{row + 1}"
    
    def supports_streaming(self) -> bool:
        """由于xlrd库限制，XLS解析器不支持流式处理。"""
        return False  # xlrd 不支持真正的流式读取，但可实现分块读取

    def create_lazy_sheet(self, file_path: str, sheet_name: str | None = None) -> LazySheet | None:
        """
        由于xlrd库限制，XLS格式不支持流式读取。
        大文件建议采用分块读取方案。
        """
        return None  # 受xlrd限制，未实现
