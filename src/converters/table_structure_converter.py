from typing import Any
from src.models.table_model import Sheet
from src.utils.range_parser import parse_range_string
from src.utils.html_utils import escape_html, create_html_element, create_table_cell


class TableStructureConverter:
    """处理 HTML 表格结构的生成。"""

    def __init__(self, cell_converter, style_converter):
        self.cell_converter = cell_converter
        self.style_converter = style_converter

    def generate_table(self, sheet: Sheet, styles: dict[str, Any], header_rows: int) -> str:
        """
        生成表格的 HTML。
        """
        occupied_cells: set[tuple[int, int]] = set()
        merged_cells_map: dict[tuple[int, int], dict[str, int]] = {}

        for merged_range in sheet.merged_cells:
            try:
                start_row, start_col, end_row, end_col = parse_range_string(merged_range)
                row_span = end_row - start_row + 1
                col_span = end_col - start_col + 1

                if row_span > 1 or col_span > 1:
                    merged_cells_map[(start_row, start_col)] = {"rowspan": row_span, "colspan": col_span}
                    for r in range(start_row, end_row + 1):
                        for c in range(start_col, end_col + 1):
                            if (r, c) != (start_row, start_col):
                                occupied_cells.add((r, c))
            except ValueError as e:
                print(f"Could not parse merged cell range '{merged_range}': {e}")

        # 创建表格开始标签，包含属性
        table_attrs = {
            'role': 'table',
            'aria-label': f'Table: {sheet.name}'
        }
        table_parts = [create_html_element('table', '', attributes=table_attrs).replace('></table>', '>')]
        
        if sheet.name and sheet.name.strip():
            caption = create_html_element('caption', f'Table: {escape_html(sheet.name)}')
            table_parts.append(caption)

        style_key_to_id_map = {self.style_converter.get_style_key(style_obj): style_id for style_id, style_obj in
                               styles.items()}

        if header_rows > 0 and len(sheet.rows) > 0:
            table_parts.append('<thead>')
            self._generate_rows_html(table_parts, sheet.rows[:header_rows], occupied_cells, merged_cells_map,
                                     style_key_to_id_map, is_header=True)
            table_parts.append('</thead>')
            if len(sheet.rows) > header_rows:
                table_parts.append('<tbody>')
                self._generate_rows_html(table_parts, sheet.rows[header_rows:], occupied_cells, merged_cells_map,
                                         style_key_to_id_map, is_header=False, row_offset=header_rows)
                table_parts.append('</tbody>')
        else:
            self._generate_rows_html(table_parts, sheet.rows, occupied_cells, merged_cells_map, style_key_to_id_map,
                                     is_header=False)

        table_parts.append('</table>')
        return "\n".join(table_parts)

    def _generate_rows_html(self, table_parts: list, rows: list, occupied_cells: set, merged_cells_map: dict,
                            style_key_to_id_map: dict, is_header: bool = False, row_offset: int = 0):
        for r_idx, row in enumerate(rows):
            actual_row_idx = r_idx + row_offset
            table_parts.append('<tr>')

            # 找到行中最后一个有内容的单元格位置
            last_content_col = self._find_last_content_column(row, actual_row_idx, occupied_cells, merged_cells_map)

            for c_idx, cell in enumerate(row.cells):
                if (actual_row_idx, c_idx) in occupied_cells:
                    continue

                # 如果超过了最后有内容的列，且当前单元格为空，则跳过
                if c_idx > last_content_col and not self._has_meaningful_content(cell):
                    continue

                style_class = ""
                css_classes = []
                if cell.style:
                    style_key = self.style_converter.get_style_key(cell.style)
                    style_id = style_key_to_id_map.get(style_key)
                    if style_id:
                        css_classes.append(style_id)
                    if cell.style.wrap_text:
                        css_classes.append("wrap-text")
                # 检查是否需要文字溢出显示（Excel特性）
                overflow_style = ""
                if self._should_overflow_text(cell, row, c_idx):
                    css_classes.append("text-overflow")
                    # 添加内联样式确保最高优先级
                    overflow_style = ' style="overflow: visible !important; white-space: nowrap !important; width: auto !important; min-width: auto !important; word-wrap: normal !important; position: relative; z-index: 5;"'

                if css_classes:
                    style_class = f' class="{" ".join(css_classes)}"'

                span_attrs = ""
                if (actual_row_idx, c_idx) in merged_cells_map:
                    spans = merged_cells_map[(actual_row_idx, c_idx)]
                    if spans["rowspan"] > 1:
                        span_attrs += f' rowspan="{spans["rowspan"]}"'
                    if spans["colspan"] > 1:
                        span_attrs += f' colspan="{spans["colspan"]}"'
                cell_html = self._generate_cell_html(cell, style_class, span_attrs, is_header, overflow_style)
                table_parts.append(cell_html)
            table_parts.append('</tr>')

    def _generate_cell_html(self, cell, style_class, span_attrs, is_header, overflow_style=""):
        cell_content = self.cell_converter.convert(cell)

        # 处理超链接
        if cell.style and cell.style.hyperlink:
            href = escape_html(cell.style.hyperlink)
            cell_content = create_html_element('a', cell_content, attributes={'href': href})

        # 构建标题属性
        title_parts = []
        if cell.style and cell.style.comment:
            title_parts.append(escape_html(cell.style.comment))
        if cell.formula:
            title_parts.append(f"Formula: {escape_html(cell.formula)}")
        
        # 构建属性字典
        cell_attrs = {}
        if title_parts:
            cell_attrs['title'] = " | ".join(title_parts)
        if cell.style and cell.style.number_format:
            cell_attrs['data-number-format'] = escape_html(cell.style.number_format)
        
        # 解析跨度属性
        rowspan = 1
        colspan = 1
        if ' rowspan="' in span_attrs:
            rowspan = int(span_attrs.split('rowspan="')[1].split('"')[0])
        if ' colspan="' in span_attrs:
            colspan = int(span_attrs.split('colspan="')[1].split('"')[0])
        
        # 解析CSS类
        css_classes = []
        if ' class="' in style_class:
            css_classes = style_class.split('class="')[1].split('"')[0].split()
        
        # 解析内联样式
        inline_styles = {}
        if overflow_style and 'style="' in overflow_style:
            style_content = overflow_style.split('style="')[1].split('"')[0]
            for style_pair in style_content.split(';'):
                if ':' in style_pair:
                    key, value = style_pair.split(':', 1)
                    inline_styles[key.strip()] = value.strip()
        
        # 使用工具函数创建表格单元格
        return create_table_cell(
            content=cell_content,
            is_header=is_header,
            rowspan=rowspan,
            colspan=colspan,
            css_classes=css_classes,
            inline_styles=inline_styles,
            title=cell_attrs.get('title', '')
        )

    def _should_overflow_text(self, cell, row, col_idx):
        """
        检查单元格是否应应用文字溢出显示（模拟Excel行为）。
        条件：
        1. 有文字内容
        2. 文字长度超过阈值
        3. 右侧单元格为空
        4. 未设置文字换行
        """
        if not cell.value:
            return False

        # 检查文字长度 (中文字符按2倍计算)
        cell_text = str(cell.value).strip()
        # 计算显示宽度：中文字符按2倍计算，英文字符按1倍计算
        display_width = sum(2 if ord(c) > 127 else 1 for c in cell_text)
        TEXT_OVERFLOW_THRESHOLD = 8  # 短文字不需要溢出的阈值
        if display_width <= TEXT_OVERFLOW_THRESHOLD:
            return False

        # 检查是否设置了文字换行
        if cell.style and cell.style.wrap_text:
            return False  # 如果设置了换行，不应该溢出

        # 检查右边的单元格是否为空
        next_col_idx = col_idx + 1
        if next_col_idx < len(row.cells):
            next_cell = row.cells[next_col_idx]
            if next_cell.value:  # 右边有内容，不应该溢出
                return False

        return True  # 满足所有条件，应该溢出显示

    def _find_last_content_column(self, row, actual_row_idx: int, occupied_cells: set, merged_cells_map: dict) -> int:
        """
        找到行中最后一个有内容的列索引。
        参数：
            row 行对象，actual_row_idx 实际行索引，occupied_cells 被占用位置集合，merged_cells_map 合并单元格映射
        返回：
            最后一个有内容的列索引，无内容则返回-1
        """
        last_content_col = -1

        for c_idx, cell in enumerate(row.cells):
            # 跳过被占用的单元格
            if (actual_row_idx, c_idx) in occupied_cells:
                continue

            # 检查是否有意义的内容
            if self._has_meaningful_content(cell):
                last_content_col = c_idx

            # 如果是合并单元格的起始位置，也算作有内容
            elif (actual_row_idx, c_idx) in merged_cells_map:
                last_content_col = c_idx

        return last_content_col

    def _has_meaningful_content(self, cell) -> bool:
        """
        检查单元格是否有有意义内容。
        参数：
            cell 单元格对象
        返回：
            有内容返回True，否则False
        """
        # 检查值
        if cell.value is not None and str(cell.value).strip():
            return True

        # 检查是否有公式
        if cell.formula:
            return True

        # 检查是否有特殊样式（背景色、边框等）
        if cell.style:
            # 如果有背景色（非默认）
            if (hasattr(cell.style, 'background_color') and
                cell.style.background_color and
                cell.style.background_color.lower() not in ['ffffff', 'white', 'none', 'auto']):
                return True

            # 如果有边框
            if (hasattr(cell.style, 'border_top') and cell.style.border_top) or \
               (hasattr(cell.style, 'border_bottom') and cell.style.border_bottom) or \
               (hasattr(cell.style, 'border_left') and cell.style.border_left) or \
               (hasattr(cell.style, 'border_right') and cell.style.border_right):
                return True

        return False
