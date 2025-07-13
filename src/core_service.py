"""
核心服务模块 - 简洁版

提供3个核心功能的业务逻辑实现：
1. 获取表格元数据
2. 解析表格数据为JSON
3. 将修改应用回文件
"""

import json
import logging
from pathlib import Path
from typing import Any, Dict, Optional

from .utils.range_parser import parse_range_string
from .parsers.factory import ParserFactory
from .models.table_model import Sheet
from .converters.html_converter import HTMLConverter
from .streaming import StreamingTableReader, ChunkFilter
from .config import config
from .cache import get_cache_manager

logger = logging.getLogger(__name__)


class CoreService:
    """核心服务类，提供表格处理的核心功能。"""
    
    def __init__(self):
        self.parser_factory = ParserFactory()
    
    def get_sheet_info(self, file_path: str) -> Dict[str, Any]:
        """
        获取表格文件的元数据信息，不加载完整数据。
        
        Args:
            file_path: 文件路径
            
        Returns:
            包含元数据的字典
        """
        try:
            # 验证文件存在
            path = Path(file_path)
            if not path.exists():
                raise FileNotFoundError(f"文件不存在: {file_path}")
            
            # 获取解析器
            parser = self.parser_factory.get_parser(file_path)
            
            # 获取解析器信息
            parser = self.parser_factory.get_parser(file_path)

            # 快速解析获取基本信息
            sheet = parser.parse(file_path)

            return {
                "file_path": str(path.absolute()),
                "file_size": path.stat().st_size,
                "file_format": path.suffix.lower(),
                "parser_type": type(parser).__name__,
                "sheet_name": sheet.name,
                "dimensions": {
                    "rows": len(sheet.rows),
                    "cols": len(sheet.rows[0].cells) if sheet.rows else 0,
                    "total_cells": sum(len(row.cells) for row in sheet.rows)
                },
                "has_merged_cells": len(sheet.merged_cells) > 0,
                "merged_cells_count": len(sheet.merged_cells),
                "supported_formats": self.parser_factory.get_supported_formats()
            }
            
        except Exception as e:
            logger.error(f"获取表格信息失败: {e}")
            raise
    
    def parse_sheet(self, file_path: str, sheet_name: Optional[str] = None, 
                   range_string: Optional[str] = None, 
                   enable_streaming: bool = True, 
                   streaming_threshold: int = config.STREAMING_THRESHOLD_CELLS) -> Dict[str, Any]:
        """
        解析表格文件为标准化的JSON格式。
        
        Args:
            file_path: 文件路径
            sheet_name: 工作表名称（可选）
            range_string: 单元格范围（可选）
            enable_streaming: 是否启用自动流式读取（默认True）
            streaming_threshold: 流式读取的单元格数量阈值（默认10000）
            
        Returns:
            标准化的TableModel JSON
        """
        try:
            # 验证文件存在
            path = Path(file_path)
            if not path.exists():
                raise FileNotFoundError(f"文件不存在: {file_path}")
            
            # 尝试从缓存获取数据
            cache_manager = get_cache_manager()
            cached_data = cache_manager.get(file_path, range_string, sheet_name)
            if cached_data is not None:
                logger.info(f"从缓存获取数据: {file_path}")
                return cached_data['data']
            
            # 获取解析器
            parser = self.parser_factory.get_parser(file_path)
            
            # 检查是否应该使用流式读取
            if enable_streaming and self._should_use_streaming(file_path, streaming_threshold):
                json_data = self._parse_sheet_streaming(file_path, sheet_name, range_string)
            else:
                # 使用传统方法
                sheet = parser.parse(file_path)
                # 转换为标准化JSON格式
                json_data = self._sheet_to_json(sheet, range_string)
            
            # 缓存解析结果
            cache_manager.set(file_path, json_data, range_string, sheet_name)
            logger.debug(f"数据已缓存: {file_path}")
            
            return json_data
            
        except Exception as e:
            logger.error(f"解析表格失败: {e}")
            raise

    def convert_to_html(self, file_path: str, output_path: Optional[str] = None,
                       page_size: Optional[int] = None, page_number: Optional[int] = None) -> Dict[str, Any]:
        """
        将表格文件转换为HTML文件。

        Args:
            file_path: 源文件路径
            output_path: 输出HTML文件路径，如果为None则生成默认路径
            page_size: 分页大小（每页行数），如果为None则不分页
            page_number: 页码（从1开始），如果为None则显示第1页

        Returns:
            转换结果信息
        """
        try:
            # 验证文件存在
            path = Path(file_path)
            if not path.exists():
                raise FileNotFoundError(f"文件不存在: {file_path}")

            # 生成默认输出路径
            if output_path is None:
                output_path = str(path.with_suffix('.html'))

            # 获取解析器并解析
            parser = self.parser_factory.get_parser(file_path)
            sheet = parser.parse(file_path)

            # 检查是否需要分页处理
            if page_size is not None and page_size > 0:
                # 使用分页HTML转换器
                from .converters.paginated_html_converter import PaginatedHTMLConverter
                html_converter = PaginatedHTMLConverter(
                    compact_mode=False,
                    page_size=page_size,
                    page_number=page_number or 1
                )
            else:
                # 使用标准HTML转换器
                html_converter = HTMLConverter(compact_mode=False)

            result = html_converter.convert_to_file(sheet, output_path)

            return result

        except Exception as e:
            logger.error(f"HTML转换失败: {e}")
            raise

    def apply_changes(self, file_path: str, table_model_json: Dict[str, Any], create_backup: bool = True) -> Dict[str, Any]:
        """
        将TableModel JSON的修改应用回原始文件。

        Args:
            file_path: 目标文件路径
            table_model_json: 包含修改的JSON数据
            create_backup: 是否创建备份文件

        Returns:
            操作结果
        """
        try:
            # 验证文件存在
            path = Path(file_path)
            if not path.exists():
                raise FileNotFoundError(f"文件不存在: {file_path}")

            # 创建备份文件（如果需要）
            backup_path = None
            if create_backup:
                backup_path = path.with_suffix(f"{path.suffix}.backup")
                import shutil
                shutil.copy2(path, backup_path)
                logger.info(f"已创建备份文件: {backup_path}")

            # 验证JSON格式
            required_fields = ["sheet_name", "headers", "rows"]
            for field in required_fields:
                if field not in table_model_json:
                    raise ValueError(f"缺少必需字段: {field}")

            # 实现真正的数据写回功能
            file_format = path.suffix.lower()
            changes_applied = 0

            if file_format == '.csv':
                changes_applied = self._write_back_csv(path, table_model_json)
            elif file_format in ['.xlsx', '.xlsm']:
                changes_applied = self._write_back_xlsx(path, table_model_json)
            elif file_format == '.xls':
                changes_applied = self._write_back_xls(path, table_model_json)
            elif file_format == '.xlsb':
                # XLSB格式比较复杂，暂时不支持写回
                raise RuntimeError("XLSB格式暂不支持数据写回，请转换为XLSX格式进行编辑")
            else:
                raise RuntimeError(f"不支持的文件格式: {file_format}")

            return {
                "status": "success",
                "message": "数据修改已成功应用",
                "file_path": str(path.absolute()),
                "backup_path": str(backup_path) if backup_path else None,
                "backup_created": create_backup,
                "changes_applied": changes_applied,
                "sheet_name": table_model_json.get("sheet_name"),
                "headers_count": len(table_model_json.get("headers", [])),
                "rows_count": len(table_model_json.get("rows", []))
            }
            
        except Exception as e:
            logger.error(f"应用修改失败: {e}")
            raise

    def _write_back_csv(self, file_path: Path, table_model_json: Dict[str, Any]) -> int:
        """
        将修改写回CSV文件。

        Args:
            file_path: CSV文件路径
            table_model_json: 包含修改的JSON数据

        Returns:
            应用的修改数量
        """
        import csv

        headers = table_model_json["headers"]
        rows = table_model_json["rows"]

        # 准备写入的数据
        csv_rows = []

        # 添加表头
        csv_rows.append(headers)

        # 添加数据行
        for row in rows:
            csv_row = []
            for cell in row:
                # 提取单元格值
                if isinstance(cell, dict) and 'value' in cell:
                    value = cell['value']
                else:
                    value = str(cell) if cell is not None else ""
                csv_row.append(str(value) if value is not None else "")
            csv_rows.append(csv_row)

        # 写入CSV文件
        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(csv_rows)

        logger.info(f"CSV文件已更新: {file_path}")
        return len(rows)  # 返回修改的行数

    def _write_back_xls(self, file_path: Path, table_model_json: Dict[str, Any]) -> int:
        """
        将修改写回XLS文件。

        Args:
            file_path: XLS文件路径
            table_model_json: 包含修改的JSON数据

        Returns:
            应用的修改数量
        """
        import xlwt

        workbook = xlwt.Workbook(encoding='utf-8')
        sheet_name = table_model_json.get("sheet_name", "Sheet1")
        worksheet = workbook.add_sheet(sheet_name)

        headers = table_model_json.get("headers", [])
        rows = table_model_json.get("rows", [])

        # 写入表头
        for col_idx, header in enumerate(headers):
            worksheet.write(0, col_idx, header)

        # 写入数据行
        changes_count = 0
        for row_idx, row_data in enumerate(rows, 1):
            for col_idx, cell_data in enumerate(row_data):
                value = cell_data.get('value') if isinstance(cell_data, dict) else cell_data
                worksheet.write(row_idx, col_idx, value)
                changes_count += 1

        workbook.save(str(file_path))
        logger.info(f"XLS文件已更新: {file_path}")
        return changes_count

    def _write_back_xlsx(self, file_path: Path, table_model_json: Dict[str, Any]) -> int:
        """
        将修改写回XLSX文件。

        Args:
            file_path: XLSX文件路径
            table_model_json: 包含修改的JSON数据

        Returns:
            应用的修改数量
        """
        import openpyxl

        headers = table_model_json["headers"]
        rows = table_model_json["rows"]

        # 打开现有的工作簿
        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active

        # 清除现有数据（保留样式）
        max_row = worksheet.max_row
        max_col = worksheet.max_column

        # 清除单元格内容但保留格式
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.value = None

        # 写入表头
        for col_idx, header in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.value = header

        # 写入数据行
        changes_count = 0
        for row_idx, row_data in enumerate(rows, 2):  # 从第2行开始（第1行是表头）
            for col_idx, cell_data in enumerate(row_data, 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)

                # 提取单元格值
                if isinstance(cell_data, dict) and 'value' in cell_data:
                    value = cell_data['value']
                else:
                    value = cell_data

                # 尝试转换数值类型
                if value is not None and value != "":
                    try:
                        # 尝试转换为数字
                        if isinstance(value, str) and value.replace('.', '').replace('-', '').isdigit():
                            if '.' in value:
                                cell.value = float(value)
                            else:
                                cell.value = int(value)
                        else:
                            cell.value = str(value)
                    except (ValueError, TypeError):
                        cell.value = str(value)
                else:
                    cell.value = None

                changes_count += 1

        # 保存工作簿
        workbook.save(file_path)
        workbook.close()

        logger.info(f"XLSX文件已更新: {file_path}")
        return changes_count

    def _sheet_to_json(self, sheet: Sheet, range_string: Optional[str] = None) -> Dict[str, Any]:
        """
        将Sheet对象转换为标准化的JSON格式。

        Args:
            sheet: Sheet对象
            range_string: 可选的范围字符串（如"A1:D10"）

        Returns:
            标准化的JSON数据
        """
        # 计算数据大小
        total_cells = self._calculate_data_size(sheet)

        # 智能大小检测
        SMALL_FILE_THRESHOLD = config.SMALL_FILE_THRESHOLD_CELLS
        LARGE_FILE_THRESHOLD = config.LARGE_FILE_THRESHOLD_CELLS

        # 处理范围选择
        if range_string:
            try:
                start_row, start_col, end_row, end_col = parse_range_string(range_string)
                return self._extract_range_data(sheet, start_row, start_col, end_row, end_col)
            except ValueError as e:
                logger.warning(f"范围解析失败: {e}, 返回完整数据")

        # 根据文件大小选择处理策略
        if total_cells >= LARGE_FILE_THRESHOLD:
            # 大文件：返回摘要
            return self._generate_summary(sheet)
        elif total_cells >= SMALL_FILE_THRESHOLD:
            # 中文件：返回完整数据+建议
            full_data = self._extract_full_data(sheet)
            full_data["size_info"] = {
                "total_cells": total_cells,
                "processing_mode": "full_with_warning",
                "recommendation": f"文件较大({total_cells}个单元格)，建议使用范围选择功能获取特定数据"
            }
            return full_data
        else:
            # 小文件：直接返回完整数据
            full_data = self._extract_full_data(sheet)
            full_data["size_info"] = {
                "total_cells": total_cells,
                "processing_mode": "full",
                "recommendation": "文件大小适中，已返回完整数据"
            }
            return full_data

    def _style_to_dict(self, style) -> Dict[str, Any]:
        """
        将Style对象转换为字典格式。

        Args:
            style: Style对象

        Returns:
            样式字典
        """
        if not style:
            return {}

        style_dict = {}

        # 字体属性
        if style.font_name:
            style_dict["font_name"] = style.font_name
        if style.font_size:
            style_dict["font_size"] = style.font_size
        if style.font_color:
            style_dict["font_color"] = style.font_color
        if style.bold:
            style_dict["bold"] = style.bold
        if style.italic:
            style_dict["italic"] = style.italic
        if style.underline:
            style_dict["underline"] = style.underline

        # 背景和对齐
        if style.background_color:
            style_dict["background_color"] = style.background_color
        if style.text_align:
            style_dict["text_align"] = style.text_align
        if style.vertical_align:
            style_dict["vertical_align"] = style.vertical_align

        # 边框
        if style.border_top:
            style_dict["border_top"] = style.border_top
        if style.border_bottom:
            style_dict["border_bottom"] = style.border_bottom
        if style.border_left:
            style_dict["border_left"] = style.border_left
        if style.border_right:
            style_dict["border_right"] = style.border_right

        # 其他属性
        if style.wrap_text:
            style_dict["wrap_text"] = style.wrap_text
        if style.number_format:
            style_dict["number_format"] = style.number_format
        if style.hyperlink:
            style_dict["hyperlink"] = style.hyperlink
        if style.comment:
            style_dict["comment"] = style.comment

        return style_dict

    def _calculate_data_size(self, sheet: Sheet) -> int:
        """计算表格的总单元格数。"""
        return sum(len(row.cells) for row in sheet.rows)


    def _extract_range_data(self, sheet: Sheet, start_row: int, start_col: int,
                           end_row: int, end_col: int) -> Dict[str, Any]:
        """提取指定范围的数据。"""
        # 验证范围有效性
        if start_row < 0 or start_col < 0:
            raise ValueError("范围起始位置不能为负数")
        if end_row >= len(sheet.rows) or start_row > end_row:
            raise ValueError(f"行范围无效: {start_row}-{end_row}, 表格只有{len(sheet.rows)}行")

        # 提取范围内的数据
        range_rows = []
        headers = []

        for row_idx in range(start_row, min(end_row + 1, len(sheet.rows))):
            row = sheet.rows[row_idx]
            row_data = []

            for col_idx in range(start_col, min(end_col + 1, len(row.cells))):
                if col_idx < len(row.cells):
                    cell = row.cells[col_idx]
                    cell_data = {
                        "value": cell.value,
                        "style": self._style_to_dict(cell.style) if cell.style else None
                    }
                else:
                    cell_data = {"value": None, "style": None}
                row_data.append(cell_data)

            if row_idx == start_row:
                # 第一行作为表头
                headers = [cell["value"] if cell["value"] is not None else f"Column_{i}"
                          for i, cell in enumerate(row_data)]
            else:
                range_rows.append(row_data)

        return {
            "sheet_name": sheet.name,
            "range": f"{chr(65 + start_col)}{start_row + 1}:{chr(65 + end_col)}{end_row + 1}",
            "metadata": {
                "parser_type": "新解析器系统",
                "range_rows": len(range_rows),
                "range_cols": end_col - start_col + 1,
                "has_styles": any(any(cell.get("style") for cell in row) for row in range_rows)
            },
            "headers": headers,
            "rows": range_rows,
            "size_info": {
                "total_cells": len(range_rows) * (end_col - start_col + 1),
                "processing_mode": "range_selection",
                "recommendation": "已返回指定范围的数据"
            }
        }

    def _generate_summary(self, sheet: Sheet) -> Dict[str, Any]:
        """为大文件生成摘要信息。"""
        total_cells = self._calculate_data_size(sheet)

        # 提取前5行作为样本
        sample_rows = []
        headers = []

        for row_idx, row in enumerate(sheet.rows[:5]):
            row_data = []
            for cell in row.cells:
                cell_data = {
                    "value": cell.value,
                    "style": self._style_to_dict(cell.style) if cell.style else None
                }
                row_data.append(cell_data)

            if row_idx == 0:
                headers = [cell["value"] if cell["value"] is not None else f"Column_{i}"
                          for i, cell in enumerate(row_data)]
            else:
                sample_rows.append(row_data)

        # 分析数据类型
        data_types = {}
        for col_idx, header in enumerate(headers):
            types_found = set()
            for row in sheet.rows[1:6]:  # 检查前5行数据
                if col_idx < len(row.cells) and row.cells[col_idx].value is not None:
                    value = row.cells[col_idx].value
                    if isinstance(value, str):
                        types_found.add("text")
                    elif isinstance(value, (int, float)):
                        types_found.add("number")
                    else:
                        types_found.add("other")
            data_types[header] = list(types_found) if types_found else ["unknown"]

        return {
            "sheet_name": sheet.name,
            "metadata": {
                "parser_type": "新解析器系统",
                "total_rows": len(sheet.rows),
                "total_cols": len(headers),
                "total_cells": total_cells,
                "has_styles": any(any(cell.style for cell in row.cells) for row in sheet.rows),
                "data_types": data_types
            },
            "sample_data": {
                "headers": headers,
                "rows": sample_rows,
                "note": f"显示前{len(sample_rows)}行数据作为样本"
            },
            "size_info": {
                "total_cells": total_cells,
                "processing_mode": "summary",
                "recommendation": f"文件过大({total_cells}个单元格)，建议使用范围选择。例如：A1:J100"
            },
            "suggested_ranges": [
                "A1:J100",  # 前100行，前10列
                "A1:Z50",   # 前50行，前26列
                f"A1:{chr(65 + min(25, len(headers) - 1))}200"  # 前200行，实际列数
            ]
        }

    def _extract_full_data(self, sheet: Sheet) -> Dict[str, Any]:
        """提取完整的表格数据。"""
        # 提取表头（假设第一行是表头）
        headers = []
        if sheet.rows:
            first_row = sheet.rows[0]
            headers = [cell.value if cell.value is not None else f"Column_{i}"
                      for i, cell in enumerate(first_row.cells)]

        # 提取数据行
        data_rows = []
        for row_idx, row in enumerate(sheet.rows[1:], 1):  # 跳过表头行
            row_data = []
            for cell in row.cells:
                cell_data = {
                    "value": cell.value,
                    "style": self._style_to_dict(cell.style) if cell.style else None
                }
                row_data.append(cell_data)
            data_rows.append(row_data)

        return {
            "sheet_name": sheet.name,
            "metadata": {
                "parser_type": "新解析器系统",
                "total_rows": len(sheet.rows),
                "total_cols": len(headers),
                "data_rows": len(data_rows),
                "has_styles": any(any(cell.style for cell in row.cells) for row in sheet.rows)
            },
            "headers": headers,
            "rows": data_rows,
            "merged_cells": sheet.merged_cells,
            "format_info": {
                "supports_styles": True,
                "supports_merged_cells": True,
                "supports_hyperlinks": True
            }
        }

    def _should_use_streaming(self, file_path: str, threshold: int) -> bool:
        """
        判断是否应该使用流式读取。
        
        Args:
            file_path: 文件路径
            threshold: 单元格数量阈值
            
        Returns:
            如果应该使用流式读取则返回True
        """
        try:
            # 检查解析器是否支持流式读取
            if not self.parser_factory.supports_streaming(file_path):
                return False
            
            # 快速估算文件大小
            path = Path(file_path)
            file_size = path.stat().st_size
            
            # 基于文件大小的简单启发式判断
            # 大于10MB的文件通常值得流式读取
            if file_size > config.STREAMING_FILE_SIZE_MB * 1024 * 1024:  # 10MB
                return True
            
            # 对于较小的文件，可以先快速解析一小部分来估算总大小
            try:
                with StreamingTableReader(file_path) as reader:
                    info = reader.get_info()
                    total_cells = info['total_rows'] * info['total_columns']
                    return total_cells >= threshold
            except Exception as e:
                logger.warning(f"无法估算文件大小: {e}")
                return False
                
        except Exception as e:
            logger.warning(f"检查流式读取支持时出错: {e}")
            return False
    
    def _parse_sheet_streaming(self, file_path: str, sheet_name: Optional[str] = None, 
                              range_string: Optional[str] = None) -> Dict[str, Any]:
        """
        使用流式读取器解析表格文件。
        
        Args:
            file_path: 文件路径
            sheet_name: 工作表名称（可选）
            range_string: 单元格范围（可选）
            
        Returns:
            标准化的TableModel JSON
        """
        try:
            with StreamingTableReader(file_path) as reader:
                # 获取文件信息
                file_info = reader.get_info()
                
                # 设置过滤配置
                filter_config = None
                if range_string:
                    filter_config = ChunkFilter(range_string=range_string)
                
                # 检查是否为大文件，如果是则返回摘要
                total_cells = file_info['total_rows'] * file_info['total_columns']
                LARGE_FILE_THRESHOLD = config.STREAMING_SUMMARY_THRESHOLD_CELLS  # 大文件阈值
                
                if total_cells > LARGE_FILE_THRESHOLD and not range_string:
                    # 返回摘要信息
                    return self._generate_streaming_summary(reader, file_info)
                
                # 读取数据块
                all_rows = []
                headers = []
                
                for chunk in reader.iter_chunks(rows=config.STREAMING_CHUNK_SIZE_ROWS, filter_config=filter_config):
                    if not headers:
                        headers = chunk.headers
                    
                    # 转换行格式
                    for row in chunk.rows:
                        row_data = []
                        for cell in row.cells:
                            cell_data = {
                                "value": cell.value,
                                "style": self._style_to_dict(cell.style) if cell.style else None
                            }
                            row_data.append(cell_data)
                        all_rows.append(row_data)
                
                # 构建返回数据
                return {
                    "sheet_name": Path(file_path).stem,
                    "metadata": {
                        "parser_type": file_info['parser_type'],
                        "total_rows": file_info['total_rows'],
                        "total_cols": file_info['total_columns'],
                        "data_rows": len(all_rows),
                        "processing_mode": "streaming",
                        "supports_streaming": True
                    },
                    "headers": headers,
                    "rows": all_rows,
                    "size_info": {
                        "total_cells": total_cells,
                        "processing_mode": "streaming",
                        "recommendation": "使用流式读取处理大文件",
                        "estimated_memory_usage": file_info['estimated_memory_usage']
                    }
                }
                
        except Exception as e:
            logger.error(f"流式解析失败: {e}")
            # 回退到传统方法
            parser = self.parser_factory.get_parser(file_path)
            sheet = parser.parse(file_path)
            return self._sheet_to_json(sheet, range_string)
    
    def _generate_streaming_summary(self, reader: StreamingTableReader, file_info: Dict[str, Any]) -> Dict[str, Any]:
        """
        为大文件生成流式摘要信息。
        
        Args:
            reader: 流式读取器
            file_info: 文件信息
            
        Returns:
            摘要数据
        """
        # 读取前几行作为样本
        sample_rows = []
        headers = []
        
        filter_config = ChunkFilter(start_row=0, max_rows=5)
        
        for chunk in reader.iter_chunks(rows=5, filter_config=filter_config):
            headers = chunk.headers
            
            for row in chunk.rows[1:]:  # 跳过表头行
                row_data = []
                for cell in row.cells:
                    cell_data = {
                        "value": cell.value,
                        "style": self._style_to_dict(cell.style) if cell.style else None
                    }
                    row_data.append(cell_data)
                sample_rows.append(row_data)
            break  # 只需要第一个块
        
        total_cells = file_info['total_rows'] * file_info['total_columns']
        
        return {
            "sheet_name": Path(reader.file_path).stem,
            "metadata": {
                "parser_type": file_info['parser_type'],
                "total_rows": file_info['total_rows'],
                "total_cols": file_info['total_columns'],
                "total_cells": total_cells,
                "processing_mode": "streaming_summary",
                "supports_streaming": True
            },
            "sample_data": {
                "headers": headers,
                "rows": sample_rows,
                "note": f"显示前{len(sample_rows)}行数据作为样本（流式读取）"
            },
            "size_info": {
                "total_cells": total_cells,
                "processing_mode": "streaming_summary",
                "recommendation": f"文件过大({total_cells}个单元格)，建议使用范围选择。例如：A1:J100",
                "estimated_memory_usage": file_info['estimated_memory_usage']
            },
            "suggested_ranges": [
                "A1:J100",  # 前100行，前10列
                "A1:Z50",   # 前50行，前26列
                f"A1:{chr(65 + min(25, len(headers) - 1))}200"  # 前200行，实际列数
            ]
        }
