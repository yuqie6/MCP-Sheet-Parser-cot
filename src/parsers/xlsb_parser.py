"""
XLSB 解析器模块

使用 pyxlsb 库解析 Excel 二进制格式 .xlsb 文件。
"""

from pyxlsb import open_workbook
from src.models.table_model import Sheet, Row, Cell, Style
from src.parsers.base_parser import BaseParser


class XlsbParser(BaseParser):
    """XLSB格式解析器，使用pyxlsb库解析Excel二进制文件。"""
    
    def parse(self, file_path: str) -> Sheet:
        """
        解析 .xlsb 文件并返回 Sheet 对象。
        
        Args:
            file_path: XLSB文件路径
            
        Returns:
            包含表格数据的Sheet对象
        """
        try:
            with open_workbook(file_path) as workbook:
                # 获取第一个工作表名称
                sheet_names = workbook.get_sheet_names()
                if not sheet_names:
                    raise ValueError("工作簿中没有找到工作表")
                
                sheet_name = sheet_names[0]
                
                rows = []
                with workbook.get_sheet(sheet_name) as worksheet:
                    # 收集所有行数据
                    all_rows = list(worksheet.rows())
                    
                    if not all_rows:
                        # 如果没有数据，返回空的Sheet
                        return Sheet(name=sheet_name, rows=[], merged_cells=[])
                    
                    # 确定最大列数
                    max_cols = 0
                    for row_data in all_rows:
                        if row_data:
                            max_col = max(cell.c for cell in row_data if cell is not None)
                            max_cols = max(max_cols, max_col + 1)
                    
                    # 按行处理数据
                    current_row = 0
                    for row_data in all_rows:
                        # 创建当前行的单元格列表
                        cells = [Cell(value=None, style=Style()) for _ in range(max_cols)]
                        
                        # 填充有数据的单元格
                        if row_data:
                            for cell_data in row_data:
                                if cell_data is not None and cell_data.c < max_cols:
                                    # 处理单元格值
                                    processed_value = self._process_cell_value(cell_data.v)
                                    
                                    # 创建基础样式（XLSB格式样式信息有限）
                                    cell_style = Style()
                                    
                                    # 更新单元格
                                    cells[cell_data.c] = Cell(
                                        value=processed_value,
                                        style=cell_style
                                    )
                        
                        rows.append(Row(cells=cells))
                        current_row += 1
                
                # XLSB格式的合并单元格信息较难获取，暂时返回空列表
                merged_cells = []
                
                return Sheet(
                    name=sheet_name,
                    rows=rows,
                    merged_cells=merged_cells
                )
                
        except Exception as e:
            raise RuntimeError(f"解析XLSB文件失败: {str(e)}")
    
    def _process_cell_value(self, value):
        """
        处理单元格值，确保类型正确。
        
        Args:
            value: 原始单元格值
            
        Returns:
            处理后的单元格值
        """
        if value is None:
            return None
        
        # 如果是字符串，直接返回
        if isinstance(value, str):
            return value
        
        # 如果是数字，检查是否为整数
        if isinstance(value, (int, float)):
            if isinstance(value, float) and value.is_integer():
                return int(value)
            return value
        
        # 如果是布尔值
        if isinstance(value, bool):
            return value
        
        # 其他类型转换为字符串
        return str(value)
    
    def _extract_basic_style(self) -> Style:
        """
        创建基础样式对象。
        
        注意：pyxlsb库对样式信息的支持有限，
        主要专注于数据提取而非格式化信息。
        
        Returns:
            基础样式对象
        """
        return Style(
            bold=False,
            italic=False,
            underline=False,
            font_color="#000000",
            background_color="#FFFFFF",
            text_align="left",
            vertical_align="top"
        )
