"""
XLSB格式解析器模块

基于pyxlsb库实现XLSB格式的解析，支持Excel二进制格式的数据提取。
XLSB格式的样式信息相对有限，主要专注于数据准确性。
"""

import logging
from datetime import datetime
from typing import Optional
from pyxlsb import open_workbook, convert_date
from src.models.table_model import Sheet, Row, Cell, Style, LazySheet
from src.parsers.base_parser import BaseParser

logger = logging.getLogger(__name__)


class XlsbParser(BaseParser):
    """XLSB格式解析器，基于pyxlsb库实现数据提取和基础样式支持。"""
    
    def parse(self, file_path: str) -> Sheet:
        """
        解析XLSB文件并返回Sheet对象。
        
        Args:
            file_path: XLSB文件路径
            
        Returns:
            包含数据和基础样式的Sheet对象
            
        Raises:
            RuntimeError: 当解析失败时
        """
        try:
            with open_workbook(file_path) as workbook:
                # 获取第一个工作表
                if not workbook.sheets:
                    raise RuntimeError("工作簿不包含任何工作表")
                
                sheet_name = workbook.sheets[0]
                
                # 打开第一个工作表
                with workbook.get_sheet(1) as worksheet:
                    rows = []
                    
                    # 读取所有行数据
                    for row_data in worksheet.rows():
                        cells = []
                        
                        # 处理当前行的所有单元格
                        if row_data:
                            # 获取最大列数
                            max_col = max(cell.c for cell in row_data) if row_data else 0
                            
                            # 创建完整的行，包括空单元格
                            for col_idx in range(max_col + 1):
                                # 查找当前列的单元格
                                cell_data = None
                                for cell in row_data:
                                    if cell.c == col_idx:
                                        cell_data = cell
                                        break
                                
                                # 获取单元格值和样式
                                if cell_data:
                                    cell_value = self._process_cell_value(cell_data.v)
                                    cell_style = self._extract_basic_style(cell_data)
                                else:
                                    cell_value = None
                                    cell_style = Style()
                                
                                # 创建Cell对象
                                cell = Cell(
                                    value=cell_value,
                                    style=cell_style
                                )
                                cells.append(cell)
                        
                        rows.append(Row(cells=cells))
                    
                    # XLSB格式中合并单元格信息较难获取，暂时返回空列表
                    merged_cells = []
                    
                    return Sheet(
                        name=sheet_name,
                        rows=rows,
                        merged_cells=merged_cells
                    )
                    
        except Exception as e:
            logger.error(f"解析XLSB文件失败: {e}")
            raise RuntimeError(f"无法解析XLSB文件 {file_path}: {str(e)}")
    
    def _process_cell_value(self, value):
        """
        处理单元格值，包括数据类型转换和日期处理。
        
        Args:
            value: 原始单元格值
            
        Returns:
            处理后的单元格值
        """
        if value is None:
            return None
        
        # 处理不同数据类型
        if isinstance(value, str):
            return value
        elif isinstance(value, (int, float)):
            # 检查是否为日期（更严格的日期范围检测）
            # Excel日期通常在1900-01-01到2099-12-31之间，对应数值范围约为1-73050
            if isinstance(value, float) and 25569 <= value <= 73050:  # 更合理的日期范围
                try:
                    # 尝试转换为日期
                    date_value = convert_date(value)
                    if isinstance(date_value, datetime):
                        # 额外验证：检查年份是否合理
                        if 1900 <= date_value.year <= 2099:
                            return date_value
                except:
                    # 如果转换失败，当作普通数字处理
                    pass
            
            # 如果是整数，返回int，否则返回float
            if isinstance(value, float) and value.is_integer():
                return int(value)
            return value
        elif isinstance(value, bool):
            return value
        else:
            # 其他类型转换为字符串
            return str(value)
    
    def _extract_basic_style(self, cell_data) -> Style:
        """
        从XLSB单元格提取基础样式信息。
        
        注意：XLSB格式的样式信息非常有限，主要提供默认样式。
        
        Args:
            cell_data: pyxlsb单元格对象
            
        Returns:
            Style对象（主要为默认值）
        """
        style = Style()
        
        # XLSB格式中样式信息获取较为困难
        # pyxlsb库主要专注于数据提取，样式支持有限
        # 这里提供基础的样式框架，可以根据需要扩展
        
        try:
            # 检查是否有基础的格式信息
            # 注意：pyxlsb库的样式支持非常有限
            # 大多数样式属性需要设置为默认值
            
            # 设置默认字体
            style.font_name = "Calibri"
            style.font_size = 11.0
            style.font_color = "#000000"
            
            # 设置默认对齐
            style.text_align = "left"
            style.vertical_align = "bottom"
            
            # 其他样式属性保持默认值（None或False）
            
        except Exception as e:
            logger.debug(f"提取XLSB样式失败: {e}")
        
        return style
    
    def _get_sheet_names(self, workbook) -> list[str]:
        """
        获取工作簿中所有工作表的名称。
        
        Args:
            workbook: pyxlsb工作簿对象
            
        Returns:
            工作表名称列表
        """
        try:
            return workbook.sheets if workbook.sheets else []
        except Exception as e:
            logger.warning(f"获取工作表名称失败: {e}")
            return []
    
    def _normalize_row_data(self, row_data, max_columns: int) -> list:
        """
        标准化行数据，确保所有行都有相同的列数。
        
        Args:
            row_data: 原始行数据
            max_columns: 最大列数
            
        Returns:
            标准化后的行数据
        """
        normalized_row = [None] * max_columns
        
        if row_data:
            for cell in row_data:
                if 0 <= cell.c < max_columns:
                    normalized_row[cell.c] = cell.v
        
        return normalized_row
    
    def supports_streaming(self) -> bool:
        """XLSB parser has limited streaming support due to pyxlsb limitations."""
        return False  # pyxlsb has some streaming capabilities but limited style support
    
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> Optional[LazySheet]:
        """
        XLSB format has limited streaming support.
        pyxlsb supports streaming but with very limited style information.
        """
        return None  # Not implemented due to limited style support in streaming mode
