"""
XLSX解析器模块

解析Excel XLSX文件并转换为Sheet对象
包含完整的样式提取、颜色处理、边框识别等功能，支持流式读取。
"""

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import Cell as OpenpyxlCell
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.chart.bar_chart import BarChart
from openpyxl.chart.line_chart import LineChart
from openpyxl.chart.pie_chart import PieChart
from openpyxl.chart.area_chart import AreaChart
from typing import Iterator,Any
from src.models.table_model import Sheet, Row, Cell, LazySheet, Chart, ChartPosition
from src.parsers.base_parser import BaseParser
from src.utils.style_parser import extract_style, extract_cell_value


class XlsxRowProvider:
    """XLSX文件的惰性行提供者，基于openpyxl的read_only流式模式。"""
    
    def __init__(self, file_path: str, sheet_name: str | None = None):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self._total_rows_cache: int | None= None
        self._merged_cells_cache: list[str] | None = None
        self._worksheet_title_cache: str | None = None
    
    def _get_worksheet_info(self):
        """无需读取全部数据即可获取工作表信息。"""
        if self._worksheet_title_cache is None:
            workbook = openpyxl.load_workbook(self.file_path, read_only=True)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            if worksheet is not None:
                self._worksheet_title_cache = worksheet.title
            else:
                self._worksheet_title_cache = ""
            workbook.close()
        return self._worksheet_title_cache
    
    def _get_merged_cells(self) -> list[str]:
        """获取合并单元格信息。"""
        if self._merged_cells_cache is None:
            # Read merged cells from non-read-only workbook (required for merged_cells access)
            workbook = openpyxl.load_workbook(self.file_path)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            if worksheet is not None and hasattr(worksheet, "merged_cells"):
                self._merged_cells_cache = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]
            else:
                self._merged_cells_cache = []
            workbook.close()
        return self._merged_cells_cache
    
    def _parse_row(self, row_cells: tuple) -> Row:
        """将openpyxl单元格元组解析为Row对象。"""
        cells = []
        for cell in row_cells:
            cell_value = cell.value
            cell_style = extract_style(cell) if cell else None
            formula = None
            if isinstance(cell, OpenpyxlCell) and hasattr(cell, 'data_type') and cell.data_type == 'f':
                if hasattr(cell, 'value') and cell.value:
                    formula = str(cell.value)
            cells.append(Cell(value=cell_value, style=cell_style, formula=formula))
        return Row(cells=cells)

    def iter_rows(self, start_row: int = 0, max_rows: int | None = None) -> Iterator[Row]:
        """按需产出完整结构的行。"""
        workbook = openpyxl.load_workbook(self.file_path, read_only=True)
        worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
        try:
            if worksheet is not None:
                # 获取工作表的完整尺寸
                max_row = worksheet.max_row or 0
                max_col = worksheet.max_column or 0

                # 计算实际的行范围
                end_row = max_row
                if max_rows is not None:
                    end_row = min(start_row + max_rows, max_row)

                # 使用坐标访问确保完整的行结构
                for row_idx in range(start_row + 1, end_row + 1):  # openpyxl使用1基索引
                    cells = []
                    for col_idx in range(1, max_col + 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cells.append(cell)
                    yield self._parse_row(tuple(cells))
        except Exception as e:
            # 确保在出现异常时也能正确关闭工作簿
            workbook.close()
            raise RuntimeError(f"流式读取XLSX文件失败: {str(e)}") from e
        finally:
            workbook.close()
    
    def get_row(self, row_index: int) -> Row:
        """按索引获取完整结构的指定行。"""
        workbook = openpyxl.load_workbook(self.file_path, read_only=True)
        worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
        try:
            if worksheet is not None:
                max_row = worksheet.max_row or 0
                max_col = worksheet.max_column or 0

                if row_index >= max_row:
                    raise IndexError(f"Row index {row_index} out of range (max: {max_row-1})")

                # 使用坐标访问获取完整的行
                cells = []
                for col_idx in range(1, max_col + 1):
                    cell = worksheet.cell(row=row_index + 1, column=col_idx)  # openpyxl使用1基索引
                    cells.append(cell)
                return self._parse_row(tuple(cells))
            raise IndexError(f"Row index {row_index} out of range")
        except IndexError:
            # 重新抛出索引错误
            workbook.close()
            raise
        except Exception as e:
            # 处理其他异常
            workbook.close()
            raise RuntimeError(f"获取XLSX行数据失败: {str(e)}") from e
        finally:
            workbook.close()
    
    def get_total_rows(self) -> int:
        """无需加载全部数据即可获取总行数。"""
        if self._total_rows_cache is None:
            workbook = openpyxl.load_workbook(self.file_path, read_only=True)
            worksheet = workbook.active if self.sheet_name is None else workbook[self.sheet_name]
            try:
                # Count rows using worksheet.max_row property
                if worksheet is not None and hasattr(worksheet, "max_row"):
                    self._total_rows_cache = worksheet.max_row or 0
                else:
                    self._total_rows_cache = 0
            finally:
                workbook.close()
        return self._total_rows_cache


class XlsxParser(BaseParser):
    """
    XLSX文件解析器，支持完整样式提取。

    该解析器处理现代Excel文件（.xlsx），可提取丰富的样式信息，包括字体、颜色、边框和数字格式。
    同时支持通过XlsxRowProvider进行大文件流式读取。
    """

    def parse(self, file_path: str) -> list[Sheet]:
        """
        解析XLSX文件，返回每个工作表对应的Sheet对象列表。
        """
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=False)
            data_only_workbook = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            raise IOError(f"无法加载Excel文件: {e}")

        sheets = []
        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            data_only_worksheet = data_only_workbook[sheet_name]
            
            # ... (rest of the parsing logic for a single sheet)
            # This part needs to be encapsulated into a helper method
            
            sheet = self._parse_sheet(worksheet, data_only_worksheet)
            sheets.append(sheet)
            
        return sheets

    def _parse_sheet(self, worksheet: Worksheet, data_only_worksheet: Worksheet) -> Sheet:
        """解析单个工作表的辅助方法。"""
        # Get sheet dimensions
        max_row = worksheet.max_row or 0
        max_col = worksheet.max_column or 0

        # Parse rows and cells
        rows = []
        for row_idx in range(1, max_row + 1):
            cells = []
            for col_idx in range(1, max_col + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                data_cell = data_only_worksheet.cell(row=row_idx, column=col_idx)
                
                cell_style = extract_style(cell)
                
                if cell.data_type == 'f' and cell.value:
                    cell_value = data_cell.value
                    formula = str(cell.value)
                else:
                    cell_value = extract_cell_value(cell)
                    formula = None
                
                parsed_cell = Cell(
                    value=cell_value,
                    style=cell_style,
                    formula=formula
                )
                cells.append(parsed_cell)
            rows.append(Row(cells=cells))

        # Extract other sheet properties
        merged_cells = [str(merged_cell_range) for merged_cell_range in worksheet.merged_cells.ranges]
        
        column_widths = {col_idx - 1: worksheet.column_dimensions[get_column_letter(col_idx)].width 
                         for col_idx in range(1, max_col + 1) if worksheet.column_dimensions[get_column_letter(col_idx)].width}
                         
        row_heights = {row_idx - 1: worksheet.row_dimensions[row_idx].height 
                       for row_idx in range(1, max_row + 1) if worksheet.row_dimensions[row_idx].height}

        default_col_width = worksheet.sheet_format.defaultColWidth or 8.43
        default_row_height = worksheet.sheet_format.defaultRowHeight or 18.0
        
        # Extract images and charts
        charts = self._extract_charts(worksheet)
        images = self._extract_images(worksheet)
        # Combine them, perhaps differentiating by type if needed in the model
        all_visuals = charts + images

        return Sheet(
            name=worksheet.title,
            rows=rows,
            merged_cells=merged_cells,
            charts=all_visuals,
            column_widths=column_widths,
            row_heights=row_heights,
            default_column_width=default_col_width,
            default_row_height=default_row_height
        )

    def _extract_images(self, worksheet: Worksheet) -> list[Chart]:
        """提取工作表中的嵌入图片。"""
        images = []
        # Use getattr to safely access _images attribute
        worksheet_images = getattr(worksheet, '_images', [])
        for image in worksheet_images:
            if isinstance(image, OpenpyxlImage):
                try:
                    # Get image data safely
                    img_data = None
                    if hasattr(image, 'ref') and image.ref:
                        img_data = image.ref
                        if not isinstance(img_data, bytes):
                            img_data = str(img_data).encode('utf-8')
                    elif hasattr(image, '_data') and getattr(image, '_data', None):
                        img_data = getattr(image, '_data', None)

                    # Get anchor position safely
                    anchor_str = "A1"  # 默认位置
                    if hasattr(image, 'anchor') and image.anchor:
                        try:
                            # 安全地处理anchor对象
                            anchor = image.anchor
                            if hasattr(anchor, '_from') and getattr(anchor, '_from', None):
                                # TwoCellAnchor类型
                                from_cell = getattr(anchor, '_from', None)
                                if from_cell and hasattr(from_cell, 'col') and hasattr(from_cell, 'row'):
                                    # 转换为Excel单元格引用
                                    from openpyxl.utils import get_column_letter
                                    col_letter = get_column_letter(from_cell.col + 1)  # openpyxl使用0基索引
                                    anchor_str = f"{col_letter}{from_cell.row + 1}"
                            else:
                                # 其他类型的anchor，转换为字符串
                                anchor_str = str(anchor)
                        except Exception:
                            anchor_str = "A1"

                    # Create a Chart object to represent the image
                    image_chart = Chart(
                        name=f"Image {len(images) + 1}",
                        type="image",
                        anchor=anchor_str
                    )
                    # Store image data in chart_data for consistency
                    image_chart.chart_data = {'image_data': img_data, 'type': 'image'}
                    images.append(image_chart)
                except Exception as e:
                    import logging
                    logger = logging.getLogger(__name__)
                    logger.error(f"Failed to extract an image: {e}")
        return images


    
    def _extract_charts(self, worksheet) -> list[Chart]:
        """提取工作表中的图表并保存原始数据。"""
        charts = []
        for chart_drawing in worksheet._charts:
            chart_type = "unknown"
            if isinstance(chart_drawing, BarChart):
                chart_type = 'bar'
            elif isinstance(chart_drawing, LineChart):
                chart_type = 'line'
            elif isinstance(chart_drawing, PieChart):
                chart_type = 'pie'
            elif isinstance(chart_drawing, AreaChart):
                chart_type = 'area'

            try:
                # 提取原始图表数据
                chart_data = self._extract_chart_data(chart_drawing, chart_type)
                
                
                # Safely get chart title
                chart_title = str(chart_drawing.title) if chart_drawing.title else f"Chart {len(charts) + 1}"
                
                # 提取详细的定位信息
                anchor_value = "A1"  # 默认位置
                position = None
                
                if hasattr(chart_drawing, 'anchor') and chart_drawing.anchor:
                    try:
                        anchor = chart_drawing.anchor
                        if hasattr(anchor, '_from') and getattr(anchor, '_from', None) and hasattr(anchor, 'to') and getattr(anchor, 'to', None):
                            # TwoCellAnchor类型 - 提取完整定位数据
                            from_cell = getattr(anchor, '_from', None)
                            to_cell = getattr(anchor, 'to', None)
                            
                            if (from_cell and to_cell and 
                                hasattr(from_cell, 'col') and hasattr(from_cell, 'row') and
                                hasattr(to_cell, 'col') and hasattr(to_cell, 'row')):
                                
                                # EMU到像素转换（96 DPI）
                                EMU_TO_PX = 96 / 914400
                                
                                # 提取位置数据
                                from_col = from_cell.col
                                from_row = from_cell.row
                                from_col_offset = getattr(from_cell, 'colOff', 0) * EMU_TO_PX
                                from_row_offset = getattr(from_cell, 'rowOff', 0) * EMU_TO_PX
                                
                                to_col = to_cell.col
                                to_row = to_cell.row  
                                to_col_offset = getattr(to_cell, 'colOff', 0) * EMU_TO_PX
                                to_row_offset = getattr(to_cell, 'rowOff', 0) * EMU_TO_PX
                                
                                # 创建详细定位对象
                                position = ChartPosition(
                                    from_col=from_col,
                                    from_row=from_row,
                                    from_col_offset=from_col_offset,
                                    from_row_offset=from_row_offset,
                                    to_col=to_col,
                                    to_row=to_row,
                                    to_col_offset=to_col_offset,
                                    to_row_offset=to_row_offset
                                )
                                
                                # 生成简单的anchor字符串用于向后兼容
                                from openpyxl.utils import get_column_letter
                                col_letter = get_column_letter(from_col + 1)  # openpyxl使用0基索引
                                anchor_value = f"{col_letter}{from_row + 1}"
                                
                        elif hasattr(anchor, 'cell'):
                            # 其他类型的anchor
                            anchor_value = str(getattr(anchor, 'cell', 'A1'))
                    except Exception as e:
                        import logging
                        logger = logging.getLogger(__name__)
                        logger.warning(f"Failed to extract positioning for chart {chart_title}: {e}")
                        anchor_value = "A1"
                        
                charts.append(Chart(
                    name=chart_title,
                    type=chart_type,
                    anchor=anchor_value,
                    chart_data=chart_data,  # 原始数据
                    position=position  # 新增：详细定位信息
                ))
            except Exception as e:
                import logging
                logger = logging.getLogger(__name__)
                logger.error(f"Failed to extract chart {chart_drawing.title}: {e}")

        return charts


    
    def supports_streaming(self) -> bool:
        """XLSX解析器支持流式处理。"""
        return True
    
    def create_lazy_sheet(self, file_path: str, sheet_name: str | None = None) -> LazySheet:
        """
        创建用于流式读取XLSX数据的LazySheet。

        参数：
            file_path: XLSX文件的绝对路径。
            sheet_name: 要解析的工作表名称（可选）。

        返回：
            可按需流式读取数据的LazySheet对象。
        """
        provider = XlsxRowProvider(file_path, sheet_name)
        name = provider._get_worksheet_info()
        merged_cells = provider._get_merged_cells()
        return LazySheet(name=name, provider=provider, merged_cells=merged_cells)

    def _extract_axis_title(self, title_obj: Any) -> str | None:
        """
        安全提取openpyxl图表轴的标题。

        参数：
            title_obj: 图表轴的标题对象

        返回：
            标题字符串，提取失败时返回None
        """
        try:
            # 多种方式尝试提取标题文本
            if hasattr(title_obj, 'tx') and title_obj.tx:
                tx = title_obj.tx
                if hasattr(tx, 'rich') and tx.rich:
                    rich = tx.rich
                    if hasattr(rich, 'p') and rich.p and len(rich.p) > 0:
                        p = rich.p[0]
                        if hasattr(p, 'r') and p.r and hasattr(p.r, 't'):
                            return str(p.r.t)

            # 回退为字符串表示
            return str(title_obj) if title_obj else None
        except Exception:
            return None
    
    def _extract_chart_data(self, chart, chart_type: str) -> dict:
        """
        提取图表的原始数据，用于SVG渲染。
        
        参数：
            chart: openpyxl图表对象
            chart_type: 图表类型
            
        返回：
            包含图表数据的字典
        """
        chart_data = {
            'type': chart_type,
            'title': str(chart.title) if chart.title else '',
            'series': [],
            'x_axis_title': '',
            'y_axis_title': '',
            'position': {},  # 添加位置信息
            'size': {},      # 添加尺寸信息
            'colors': []     # 添加原始颜色信息
        }
        
        # 提取位置和尺寸信息
        if hasattr(chart, 'anchor') and chart.anchor:
            try:
                anchor = chart.anchor
                if hasattr(anchor, '_from') and getattr(anchor, '_from', None):
                    from_cell = getattr(anchor, '_from', None)
                    if from_cell:
                        chart_data['position']['from_col'] = getattr(from_cell, 'col', 0)
                        chart_data['position']['from_row'] = getattr(from_cell, 'row', 0)
                        chart_data['position']['from_col_offset'] = getattr(from_cell, 'colOff', 0)
                        chart_data['position']['from_row_offset'] = getattr(from_cell, 'rowOff', 0)
                
                if hasattr(anchor, 'to') and getattr(anchor, 'to', None):
                    to_cell = getattr(anchor, 'to', None)
                    if to_cell:
                        chart_data['position']['to_col'] = getattr(to_cell, 'col', 0)
                        chart_data['position']['to_row'] = getattr(to_cell, 'row', 0)
                        chart_data['position']['to_col_offset'] = getattr(to_cell, 'colOff', 0)
                        chart_data['position']['to_row_offset'] = getattr(to_cell, 'rowOff', 0)
                
                if hasattr(anchor, 'ext') and anchor.ext:
                    ext = anchor.ext
                    chart_data['size']['width_emu'] = ext.cx
                    chart_data['size']['height_emu'] = ext.cy
                    # EMU转像素 (1 EMU = 1/914400 inch, 1 inch = 96 px)
                    chart_data['size']['width_px'] = int(ext.cx / 914400 * 96)
                    chart_data['size']['height_px'] = int(ext.cy / 914400 * 96)
            except Exception:
                pass
        
        # 提取轴标题
        try:
            if chart.x_axis and chart.x_axis.title:
                x_title = self._extract_axis_title(chart.x_axis.title)
                if x_title:
                    chart_data['x_axis_title'] = x_title
            if chart.y_axis and chart.y_axis.title:
                y_title = self._extract_axis_title(chart.y_axis.title)
                if y_title:
                    chart_data['y_axis_title'] = y_title
        except Exception:
            pass
        
        # 提取系列数据
        if isinstance(chart, (BarChart, LineChart, AreaChart)):
            for series in chart.series:
                series_data = {
                    'name': series.tx.v if series.tx else f"Series {len(chart_data['series']) + 1}",
                    'x_data': [],
                    'y_data': [],
                    'color': None  # 添加颜色信息
                }
                
                # 提取系列颜色
                series_color = self._extract_series_color(series)
                if series_color:
                    series_data['color'] = series_color
                    chart_data['colors'].append(series_color)
                
                # 获取y轴数据
                y_data = self._extract_series_y_data(series)
                if y_data:
                    series_data['y_data'] = y_data
                
                # 获取x轴数据
                x_data = self._extract_series_x_data(series)
                if x_data:
                    series_data['x_data'] = x_data
                elif y_data:
                    # 如果没有x轴数据，生成默认标签
                    series_data['x_data'] = [f"Item {i+1}" for i in range(len(y_data))]
                
                if series_data['y_data']:
                    chart_data['series'].append(series_data)
                    
        elif isinstance(chart, PieChart):
            if chart.series and len(chart.series) > 0:
                series = chart.series[0]
                series_data = {
                    'name': series.tx.v if series.tx else "Pie Series",
                    'x_data': [],  # 标签
                    'y_data': [],   # 数值
                    'colors': []   # 饼图每个片段的颜色
                }
                
                # 获取标签数据
                x_data = self._extract_series_x_data(series)
                if x_data:
                    series_data['x_data'] = x_data
                
                # 获取数值数据
                y_data = self._extract_series_y_data(series)
                if y_data:
                    series_data['y_data'] = y_data
                
                # 提取饼图颜色（每个数据点可能有不同颜色）
                pie_colors = self._extract_pie_chart_colors(series)
                if pie_colors:
                    series_data['colors'] = pie_colors
                    chart_data['colors'] = pie_colors
                
                # 如果没有标签，生成默认标签
                if not series_data['x_data'] and series_data['y_data']:
                    series_data['x_data'] = [f"Item {i+1}" for i in range(len(series_data['y_data']))]
                
                if series_data['y_data']:
                    chart_data['series'].append(series_data)
        
        return chart_data
    
    def _extract_series_y_data(self, series) -> list:
        """提取系列的Y轴数据。"""
        y_data = []
        
        # 方法1：从series.val获取数值数据
        if hasattr(series, 'val') and series.val:
            if hasattr(series.val, 'numRef') and series.val.numRef:
                try:
                    if hasattr(series.val.numRef, 'numCache') and series.val.numRef.numCache:
                        # 从缓存中获取数据
                        y_data = [float(pt.v) for pt in series.val.numRef.numCache.pt if pt.v is not None]
                    else:
                        # 从引用中获取数据
                        y_data = [float(p.v) for p in series.val.numRef.get_rows() if p.v is not None]
                except (AttributeError, TypeError, ValueError):
                    pass

        # 方法2：从series.yVal获取（备用方法）
        if not y_data and hasattr(series, 'yVal') and series.yVal and hasattr(series.yVal, 'numRef') and series.yVal.numRef:
            try:
                y_data = [float(p.v) for p in series.yVal.numRef.get_rows() if p.v is not None]
            except (AttributeError, TypeError, ValueError):
                pass
        
        return y_data
    
    def _extract_series_x_data(self, series) -> list:
        """提取系列的X轴数据。"""
        x_data = []
        
        # 方法1：从series.cat获取分类数据
        if hasattr(series, 'cat') and series.cat:
            if hasattr(series.cat, 'strRef') and series.cat.strRef:
                try:
                    if hasattr(series.cat.strRef, 'strCache') and series.cat.strRef.strCache:
                        # 从缓存中获取数据
                        x_data = [str(pt.v) for pt in series.cat.strRef.strCache.pt if pt.v is not None]
                    else:
                        # 从引用中获取数据
                        x_data = [str(p.v) for p in series.cat.strRef.get_rows() if p.v is not None]
                except (AttributeError, TypeError):
                    pass
            elif hasattr(series.cat, 'numRef') and series.cat.numRef:
                try:
                    if hasattr(series.cat.numRef, 'numCache') and series.cat.numRef.numCache:
                        x_data = [str(pt.v) for pt in series.cat.numRef.numCache.pt if pt.v is not None]
                    else:
                        x_data = [str(p.v) for p in series.cat.numRef.get_rows() if p.v is not None]
                except (AttributeError, TypeError):
                    pass

        # 方法2：从series.xVal获取（备用方法）
        if not x_data and hasattr(series, 'xVal') and series.xVal:
            if hasattr(series.xVal, 'numRef') and series.xVal.numRef:
                try:
                    x_data = [str(p.v) for p in series.xVal.numRef.get_rows() if p.v is not None]
                except (AttributeError, TypeError):
                    pass
            elif hasattr(series.xVal, 'strRef') and series.xVal.strRef:
                try:
                    x_data = [str(p.v) for p in series.xVal.strRef.get_rows() if p.v is not None]
                except (AttributeError, TypeError):
                    pass
        
        return x_data
    
    def _extract_series_color(self, series) -> str | None:
        """
        提取系列的颜色信息。
        
        参数：
            series: openpyxl图表系列对象
            
        返回：
            颜色的十六进制字符串，提取失败时返回None
        """
        try:
            # 尝试从图形属性中获取颜色
            if hasattr(series, 'graphicalProperties') and series.graphicalProperties:
                graphic_props = series.graphicalProperties
                
                # 检查solidFill
                if hasattr(graphic_props, 'solidFill') and graphic_props.solidFill:
                    solid_fill = graphic_props.solidFill
                    
                    # RGB颜色
                    if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr:
                        rgb = solid_fill.srgbClr.val
                        if rgb:
                            return f"#{rgb.upper()}"
                    
                    # 主题颜色
                    if hasattr(solid_fill, 'schemeClr') and solid_fill.schemeClr:
                        scheme_color = solid_fill.schemeClr.val
                        return self._convert_scheme_color_to_hex(scheme_color)
                
                # 检查线条属性中的颜色
                if hasattr(graphic_props, 'ln') and graphic_props.ln:
                    line = graphic_props.ln
                    if hasattr(line, 'solidFill') and line.solidFill:
                        solid_fill = line.solidFill
                        if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr:
                            rgb = solid_fill.srgbClr.val
                            if rgb:
                                return f"#{rgb.upper()}"
            
            # 尝试从spPr属性获取
            if hasattr(series, 'spPr') and series.spPr:
                sp_pr = series.spPr
                if hasattr(sp_pr, 'solidFill') and sp_pr.solidFill:
                    solid_fill = sp_pr.solidFill
                    if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr:
                        rgb = solid_fill.srgbClr.val
                        if rgb:
                            return f"#{rgb.upper()}"
                    if hasattr(solid_fill, 'schemeClr') and solid_fill.schemeClr:
                        scheme_color = solid_fill.schemeClr.val
                        return self._convert_scheme_color_to_hex(scheme_color)
            
            return None
            
        except Exception:
            return None
    
    def _extract_pie_chart_colors(self, series) -> list[str]:
        """
        提取饼图各片段的颜色。
        
        参数：
            series: openpyxl饼图系列对象
            
        返回：
            颜色列表
        """
        colors = []
        try:
            # 检查数据点的个别颜色设置
            if hasattr(series, 'dPt') and series.dPt:
                for data_point in series.dPt:
                    color = self._extract_data_point_color(data_point)
                    if color:
                        colors.append(color)
            
            # 如果没有找到个别颜色，使用系列默认颜色
            if not colors:
                series_color = self._extract_series_color(series)
                if series_color:
                    # 生成基于主色的渐变色
                    colors = self._generate_pie_color_variants(series_color, len(series.val.numRef.numCache.pt) if hasattr(series, 'val') and series.val and hasattr(series.val, 'numRef') and series.val.numRef and hasattr(series.val.numRef, 'numCache') else 3)
            
        except Exception:
            pass
        
        return colors
    
    def _extract_data_point_color(self, data_point) -> str | None:
        """
        提取数据点的颜色。
        
        参数：
            data_point: 数据点对象
            
        返回：
            颜色的十六进制字符串
        """
        try:
            if hasattr(data_point, 'spPr') and data_point.spPr:
                sp_pr = data_point.spPr
                if hasattr(sp_pr, 'solidFill') and sp_pr.solidFill:
                    solid_fill = sp_pr.solidFill
                    if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr:
                        rgb = solid_fill.srgbClr.val
                        if rgb:
                            return f"#{rgb.upper()}"
                    if hasattr(solid_fill, 'schemeClr') and solid_fill.schemeClr:
                        scheme_color = solid_fill.schemeClr.val
                        return self._convert_scheme_color_to_hex(scheme_color)
            return None
        except Exception:
            return None
    
    def _convert_scheme_color_to_hex(self, scheme_color: str) -> str:
        """
        将Excel主题颜色转换为十六进制颜色。
        
        参数：
            scheme_color: Excel主题颜色名称
            
        返回：
            十六进制颜色字符串
        """
        # Excel主题颜色映射
        excel_theme_colors = {
            'accent1': '#5B9BD5',  # 蓝色
            'accent2': '#70AD47',  # 绿色  
            'accent3': '#FFC000',  # 橙色
            'accent4': '#E15759',  # 红色
            'accent5': '#4472C4',  # 深蓝色
            'accent6': '#FF6B35',  # 橙红色
            'dk1': '#000000',      # 深色1（黑色）
            'lt1': '#FFFFFF',      # 浅色1（白色）
            'dk2': '#44546A',      # 深色2（深灰蓝）
            'lt2': '#E7E6E6',      # 浅色2（浅灰）
            'bg1': '#FFFFFF',      # 背景1
            'bg2': '#E7E6E6',      # 背景2
            'tx1': '#000000',      # 文本1
            'tx2': '#44546A',      # 文本2
        }
        
        return excel_theme_colors.get(scheme_color, '#5B9BD5')  # 默认蓝色
    
    def _generate_pie_color_variants(self, base_color: str, count: int) -> list[str]:
        """
        基于基础颜色生成饼图的颜色变体。
        
        参数：
            base_color: 基础颜色（十六进制）
            count: 需要的颜色数量
            
        返回：
            颜色列表
        """
        if count <= 1:
            return [base_color]
        
        # 从基础颜色生成HSL变体
        try:
            # 简单的颜色变体算法：调整亮度和饱和度
            base_rgb = base_color.lstrip('#')
            r = int(base_rgb[:2], 16)
            g = int(base_rgb[2:4], 16)
            b = int(base_rgb[4:6], 16)
            
            colors = [base_color]
            
            for i in range(1, count):
                # 调整亮度
                factor = 0.8 + (i * 0.4 / count)  # 0.8 到 1.2
                new_r = min(255, int(r * factor))
                new_g = min(255, int(g * factor))
                new_b = min(255, int(b * factor))
                
                new_color = f"#{new_r:02X}{new_g:02X}{new_b:02X}"
                colors.append(new_color)
            
            return colors
            
        except Exception:
            # 如果转换失败，返回默认颜色序列
            default_colors = ['#5B9BD5', '#70AD47', '#FFC000', '#E15759', '#4472C4', '#FF6B35']
            return (default_colors * ((count // len(default_colors)) + 1))[:count]
