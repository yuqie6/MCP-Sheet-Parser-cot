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
import io
import matplotlib.pyplot as plt
from src.models.table_model import Sheet, Row, Cell, LazySheet, Chart
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
            charts=all_visuals, # Use the combined list
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
                    # Get image data - use a simple approach for type safety
                    img_data = getattr(image, 'ref', None)
                    if img_data and not isinstance(img_data, bytes):
                        img_data = str(img_data).encode('utf-8')

                    # Get anchor position safely - use string representation
                    anchor_str = str(getattr(image, 'anchor', 'A1'))

                    # Create a Chart object to represent the image
                    image_chart = Chart(
                        name=f"Image {len(images) + 1}",
                        type="image",
                        image_data=img_data,
                        anchor=anchor_str
                    )
                    images.append(image_chart)
                except Exception as e:
                    import logging
                    logger = logging.getLogger(__name__)
                    logger.error(f"Failed to extract an image: {e}")
        return images


    
    def _extract_charts(self, worksheet) -> list[Chart]:
        """提取工作表中的图表并渲染为图片。"""
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
                img_data = self._render_chart(chart_drawing)
                # Safely get chart title
                chart_title = str(chart_drawing.title) if chart_drawing.title else f"Chart {len(charts) + 1}"
                # Safely get anchor
                anchor_value = str(getattr(chart_drawing.anchor, 'cell', 'A1'))
                charts.append(Chart(
                    name=chart_title,
                    type=chart_type,
                    image_data=img_data,
                    anchor=anchor_value
                ))
            except Exception as e:
                import logging
                logger = logging.getLogger(__name__)
                logger.error(f"Failed to render chart {chart_drawing.title}: {e}")

        return charts

    def _render_chart(self, chart) ->bytes | None:
        """使用matplotlib将图表渲染为图片。"""
        fig, ax = plt.subplots()

        try:
            if isinstance(chart, (BarChart, LineChart, AreaChart)):
                for series in chart.series:
                    # 确保xVal和yVal存在且包含数据
                    if series.xVal and series.xVal.numRef:
                        x = [p.v for p in series.xVal.numRef.get_rows()]
                    elif series.xVal and series.xVal.strRef:
                        x = [str(p.v) for p in series.xVal.strRef.get_rows()]
                    else:
                        # 如果没有x轴数据，则使用默认的序列
                        x = range(len(series.yVal.numRef.get_rows()))

                    if series.yVal and series.yVal.numRef:
                        y = [p.v for p in series.yVal.numRef.get_rows()]
                    else:
                        # 没有y轴数据，无法绘制
                        continue
                    
                    series_label = series.tx.v if series.tx else f"Series {len(ax.lines) + 1}"

                    if isinstance(chart, BarChart):
                        ax.bar(x, y, label=series_label)
                    elif isinstance(chart, LineChart):
                        ax.plot(x, y, label=series_label)
                    elif isinstance(chart, AreaChart):
                        ax.fill_between(x, y, label=series_label, alpha=0.4)

                ax.set_title(str(chart.title) if chart.title else "Chart")
                # Safely extract axis titles
                if chart.x_axis and chart.x_axis.title:
                    x_title = self._extract_axis_title(chart.x_axis.title)
                    if x_title:
                        ax.set_xlabel(x_title)
                if chart.y_axis and chart.y_axis.title:
                    y_title = self._extract_axis_title(chart.y_axis.title)
                    if y_title:
                        ax.set_ylabel(y_title)
                ax.legend()

            elif isinstance(chart, PieChart):
                if chart.series and chart.series[0].xVal and chart.series[0].yVal:
                    labels = [p.v for p in chart.series[0].xVal.strRef.get_rows()]
                    sizes = [p.v for p in chart.series[0].yVal.numRef.get_rows()]
                    ax.pie(sizes, labels=labels, autopct='%1.1f%%')
                    ax.set_title(str(chart.title) if chart.title else "Chart")
                else:
                    raise ValueError("Pie chart data is missing or invalid.")

            else:
                # 不支持的图表类型
                plt.close(fig)
                return None

            buf = io.BytesIO()
            plt.savefig(buf, format='png')
            plt.close(fig)
            buf.seek(0)
            return buf.getvalue()
        
        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.error(f"Failed to render chart '{chart.title}': {e}")
            plt.close(fig)  # 确保在异常时关闭图形
            return None

    
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
