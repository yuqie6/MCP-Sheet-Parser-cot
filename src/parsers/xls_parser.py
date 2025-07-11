"""
XLS 解析器模块

使用 xlrd 库解析传统的 .xls 格式文件。
"""

import xlrd
from src.models.table_model import Sheet, Row, Cell, Style
from src.parsers.base_parser import BaseParser


class XlsParser(BaseParser):
    """XLS格式解析器，使用xlrd库解析传统Excel文件。"""
    
    def parse(self, file_path: str) -> Sheet:
        """
        解析 .xls 文件并返回 Sheet 对象。
        
        Args:
            file_path: XLS文件路径
            
        Returns:
            包含表格数据和样式的Sheet对象
        """
        try:
            # 打开工作簿
            workbook = xlrd.open_workbook(file_path, formatting_info=True)
            worksheet = workbook.sheet_by_index(0)  # 获取第一个工作表
            
            rows = []
            for row_idx in range(worksheet.nrows):
                cells = []
                for col_idx in range(worksheet.ncols):
                    # 获取单元格值
                    cell_value = worksheet.cell_value(row_idx, col_idx)
                    cell_type = worksheet.cell_type(row_idx, col_idx)
                    
                    # 处理不同类型的单元格值
                    processed_value = self._process_cell_value(cell_value, cell_type, workbook)
                    
                    # 提取样式信息
                    cell_style = self._extract_style(workbook, worksheet, row_idx, col_idx)
                    
                    # 创建单元格对象
                    parsed_cell = Cell(
                        value=processed_value,
                        style=cell_style
                    )
                    cells.append(parsed_cell)
                
                rows.append(Row(cells=cells))
            
            # XLS格式的合并单元格信息较难获取，暂时返回空列表
            merged_cells = []
            
            return Sheet(
                name=worksheet.name,
                rows=rows,
                merged_cells=merged_cells
            )
            
        except Exception as e:
            raise RuntimeError(f"解析XLS文件失败: {str(e)}")
    
    def _process_cell_value(self, value, cell_type, workbook):
        """
        处理不同类型的单元格值。
        
        Args:
            value: 原始单元格值
            cell_type: 单元格类型
            workbook: 工作簿对象
            
        Returns:
            处理后的单元格值
        """
        # xlrd 单元格类型常量
        XL_CELL_EMPTY = 0
        XL_CELL_TEXT = 1
        XL_CELL_NUMBER = 2
        XL_CELL_DATE = 3
        XL_CELL_BOOLEAN = 4
        XL_CELL_ERROR = 5
        
        if cell_type == XL_CELL_EMPTY:
            return None
        elif cell_type == XL_CELL_TEXT:
            return str(value)
        elif cell_type == XL_CELL_NUMBER:
            # 检查是否为日期值（使用合理的日期范围判断）
            try:
                # Excel日期值通常在1到100000之间
                if workbook and hasattr(workbook, 'datemode') and 1 <= value <= 100000:
                    try:
                        # 尝试转换为日期元组
                        date_tuple = xlrd.xldate.xldate_as_tuple(value, workbook.datemode)
                        # 检查是否为合理的日期（年份大于1900）
                        if date_tuple[0] > 1900 and date_tuple[1] >= 1 and date_tuple[2] >= 1:
                            return f"{date_tuple[0]}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
                    except (xlrd.xldate.XLDateError, ValueError):
                        # 如果日期转换失败，按数字处理
                        pass
            except:
                pass

            # 如果是整数，返回整数；否则返回浮点数
            return int(value) if value.is_integer() else value
        elif cell_type == XL_CELL_BOOLEAN:
            return bool(value)
        elif cell_type == XL_CELL_ERROR:
            return f"#ERROR:{value}"
        else:
            return value
    
    def _extract_style(self, workbook, worksheet, row_idx, col_idx) -> Style:
        """
        从XLS文件中提取样式信息。
        
        Args:
            workbook: 工作簿对象
            worksheet: 工作表对象
            row_idx: 行索引
            col_idx: 列索引
            
        Returns:
            样式对象
        """
        style = Style()
        
        try:
            # 获取单元格的格式信息
            cell_xf_index = worksheet.cell_xf_index(row_idx, col_idx)
            cell_xf = workbook.xf_list[cell_xf_index]
            
            # 提取字体信息
            if cell_xf.font_index < len(workbook.font_list):
                font = workbook.font_list[cell_xf.font_index]
                
                # 字体样式
                style.bold = bool(font.bold)
                style.italic = bool(font.italic)
                style.underline = bool(font.underline_type)
                
                # 字体大小（xlrd中字体大小单位是twips，需要转换）
                if hasattr(font, 'height'):
                    style.font_size = font.height / 20.0  # 转换为点数
                
                # 字体名称
                if hasattr(font, 'name'):
                    style.font_name = font.name
                
                # 字体颜色（简化处理）
                if hasattr(font, 'colour_index') and font.colour_index:
                    style.font_color = self._get_color_from_index(workbook, font.colour_index)
            
            # 提取背景色信息
            if hasattr(cell_xf, 'background') and cell_xf.background:
                bg_color_index = cell_xf.background.background_colour_index
                if bg_color_index and bg_color_index != 64:  # 64是自动颜色
                    style.background_color = self._get_color_from_index(workbook, bg_color_index)
            
            # 提取对齐信息
            if hasattr(cell_xf, 'alignment'):
                alignment = cell_xf.alignment
                
                # 水平对齐
                h_align_map = {0: 'left', 1: 'left', 2: 'center', 3: 'right', 4: 'justify'}
                if hasattr(alignment, 'hor_align'):
                    style.text_align = h_align_map.get(alignment.hor_align, 'left')
                
                # 垂直对齐
                v_align_map = {0: 'top', 1: 'middle', 2: 'bottom', 3: 'justify'}
                if hasattr(alignment, 'vert_align'):
                    style.vertical_align = v_align_map.get(alignment.vert_align, 'top')
                
                # 文本换行
                if hasattr(alignment, 'wrap'):
                    style.wrap_text = bool(alignment.wrap)
            
            # 边框信息（XLS格式的边框信息较复杂，暂时简化处理）
            if hasattr(cell_xf, 'border'):
                border = cell_xf.border
                if hasattr(border, 'top_line_style') and border.top_line_style:
                    style.border_top = "1px solid"
                if hasattr(border, 'bottom_line_style') and border.bottom_line_style:
                    style.border_bottom = "1px solid"
                if hasattr(border, 'left_line_style') and border.left_line_style:
                    style.border_left = "1px solid"
                if hasattr(border, 'right_line_style') and border.right_line_style:
                    style.border_right = "1px solid"
            
        except Exception:
            # 如果样式提取失败，返回默认样式
            pass
        
        return style
    
    def _get_color_from_index(self, workbook, color_index) -> str:
        """
        从颜色索引获取十六进制颜色值。
        
        Args:
            workbook: 工作簿对象
            color_index: 颜色索引
            
        Returns:
            十六进制颜色值
        """
        # 标准Excel颜色索引映射
        standard_colors = {
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
            11: "#00FF00",  # 绿色
            12: "#0000FF",  # 蓝色
            13: "#FFFF00",  # 黄色
            14: "#FF00FF",  # 洋红
            15: "#00FFFF",  # 青色
            16: "#800000",  # 深红
            17: "#008000",  # 深绿
            18: "#000080",  # 深蓝
            19: "#808000",  # 橄榄色
            20: "#800080",  # 紫色
            21: "#008080",  # 青绿色
            22: "#C0C0C0",  # 银色
            23: "#808080",  # 灰色
            64: "#000000",  # 自动颜色（默认黑色）
        }
        
        # 尝试从工作簿的颜色表中获取
        try:
            if hasattr(workbook, 'colour_map') and color_index in workbook.colour_map:
                rgb = workbook.colour_map[color_index]
                if rgb:
                    return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        except:
            pass
        
        # 使用标准颜色映射
        return standard_colors.get(color_index, "#000000")
