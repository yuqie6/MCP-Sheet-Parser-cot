"""
XLSM格式解析器模块

基于openpyxl库实现XLSM格式（Excel宏文件）的解析。
XLSM本质上是带宏的XLSX，样式提取与XLSX完全相同。
"""

import logging
import openpyxl
from typing import Optional
from src.models.table_model import Sheet, Row, Cell, Style, LazySheet
from src.parsers.xlsx_parser import XlsxParser

logger = logging.getLogger(__name__)


class XlsmParser(XlsxParser):
    """
    XLSM格式解析器，继承XlsxParser实现。
    
    XLSM格式与XLSX在数据和样式方面完全相同，主要区别是包含VBA宏代码。
    本解析器保留宏信息但不执行，专注于数据和样式提取。
    """
    
    def parse(self, file_path: str) -> Sheet:
        """
        解析XLSM文件并返回Sheet对象。
        
        Args:
            file_path: XLSM文件路径
            
        Returns:
            包含完整数据和样式的Sheet对象
            
        Raises:
            RuntimeError: 当解析失败时
        """
        try:
            # 使用keep_vba=True保留宏信息，但不执行
            workbook = openpyxl.load_workbook(file_path, keep_vba=True)
            worksheet = workbook.active

            if worksheet is None:
                raise ValueError("工作簿不包含任何活动工作表")
            
            # 记录宏文件信息（用于调试和日志）
            self._log_macro_info(workbook, file_path)
            
            # 解析数据和样式（完全复用XlsxParser的逻辑）
            rows = []
            for row in worksheet.iter_rows():
                cells = []
                for cell in row:
                    # 提取单元格值和样式
                    cell_value = cell.value
                    cell_style = self._extract_style(cell)

                    # 创建Cell对象
                    parsed_cell = Cell(
                        value=cell_value,
                        style=cell_style
                    )
                    cells.append(parsed_cell)
                rows.append(Row(cells=cells))
                
            # 提取合并单元格信息
            merged_cells = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]

            return Sheet(
                name=worksheet.title,
                rows=rows,
                merged_cells=merged_cells
            )
            
        except Exception as e:
            logger.error(f"解析XLSM文件失败: {e}")
            raise RuntimeError(f"无法解析XLSM文件 {file_path}: {str(e)}")
    
    def _log_macro_info(self, workbook, file_path: str) -> None:
        """
        记录宏文件的相关信息，用于调试和日志。
        
        Args:
            workbook: openpyxl工作簿对象
            file_path: 文件路径
        """
        try:
            # 检查是否包含VBA项目
            has_vba = hasattr(workbook, 'vba_archive') and workbook.vba_archive is not None
            
            if has_vba:
                logger.info(f"XLSM文件 {file_path} 包含VBA宏代码，已保留但不执行")
                
                # 尝试获取VBA项目的基本信息
                try:
                    vba_archive = workbook.vba_archive
                    if hasattr(vba_archive, 'namelist'):
                        vba_files = vba_archive.namelist()
                        logger.debug(f"VBA文件数量: {len(vba_files)}")
                        
                        # 记录主要的VBA组件
                        vba_modules = [f for f in vba_files if f.endswith('.bas') or f.endswith('.cls') or f.endswith('.frm')]
                        if vba_modules:
                            logger.debug(f"VBA模块: {vba_modules}")
                            
                except Exception as vba_e:
                    logger.debug(f"获取VBA详细信息失败: {vba_e}")
            else:
                logger.info(f"XLSM文件 {file_path} 不包含VBA宏代码")
                
        except Exception as e:
            logger.debug(f"记录宏信息失败: {e}")
    
    def get_macro_info(self, file_path: str) -> dict:
        """
        获取XLSM文件的宏信息摘要。
        
        Args:
            file_path: XLSM文件路径
            
        Returns:
            包含宏信息的字典
        """
        macro_info = {
            "has_macros": False,
            "vba_files_count": 0,
            "vba_modules": [],
            "file_path": file_path
        }
        
        try:
            workbook = openpyxl.load_workbook(file_path, keep_vba=True)
            
            # 检查VBA存在性
            if hasattr(workbook, 'vba_archive') and workbook.vba_archive is not None:
                macro_info["has_macros"] = True
                
                try:
                    vba_archive = workbook.vba_archive
                    if hasattr(vba_archive, 'namelist'):
                        vba_files = vba_archive.namelist()
                        macro_info["vba_files_count"] = len(vba_files)
                        
                        # 提取VBA模块名称
                        vba_modules = [f for f in vba_files if f.endswith('.bas') or f.endswith('.cls') or f.endswith('.frm')]
                        macro_info["vba_modules"] = vba_modules
                        
                except Exception as vba_e:
                    logger.debug(f"获取VBA详细信息失败: {vba_e}")
            
            workbook.close()
            
        except Exception as e:
            logger.warning(f"获取宏信息失败: {e}")
        
        return macro_info
    
    def supports_streaming(self) -> bool:
        """XLSM parser supports streaming (inherits from XLSX)."""
        return True
    
    def create_lazy_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> LazySheet:
        """
        Create a lazy sheet for XLSM that can stream data on demand.
        XLSM files are processed the same as XLSX files for streaming.
        
        Args:
            file_path: XLSM文件路径
            sheet_name: 工作表名称（可选）
            
        Returns:
            LazySheet对象
        """
        # Use the same streaming implementation as XLSX
        from .xlsx_parser import XlsxRowProvider
        
        provider = XlsxRowProvider(file_path, sheet_name)
        name = provider._get_worksheet_info()
        merged_cells = provider._get_merged_cells()
        return LazySheet(name=name, provider=provider, merged_cells=merged_cells)
    
    def is_macro_enabled_file(self, file_path: str) -> bool:
        """
        检查文件是否为启用宏的Excel文件。
        
        Args:
            file_path: 文件路径
            
        Returns:
            如果文件包含宏则返回True，否则返回False
        """
        try:
            # 首先检查文件扩展名
            if not file_path.lower().endswith('.xlsm'):
                return False
            
            # 然后检查实际的宏内容
            macro_info = self.get_macro_info(file_path)
            return macro_info["has_macros"]
            
        except Exception as e:
            logger.debug(f"检查宏状态失败: {e}")
            return False
