from src.models.table_model import Sheet
from src.utils.range_parser import parse_range_string
from src.utils.html_utils import escape_html


class TableStructureConverter:
    """Handles the generation of the HTML table structure."""

    def __init__(self, cell_converter, style_converter):
        self.cell_converter = cell_converter
        self.style_converter = style_converter

    def generate_table(self, sheet: Sheet, styles: dict[str, any], header_rows: int) -> str:
        """
        Generates the HTML for a table.
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

        table_parts = [f'<table role="table" aria-label="Table: {sheet.name}">']
        if sheet.name and sheet.name.strip():
            table_parts.append(f'<caption>Table: {escape_html(sheet.name)}</caption>')

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
            for c_idx, cell in enumerate(row.cells):
                if (actual_row_idx, c_idx) in occupied_cells:
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

        if cell.style and cell.style.hyperlink:
            href = escape_html(cell.style.hyperlink)
            cell_content = f'<a href="{href}">{cell_content}</a>'

        title_attr = ""
        if cell.style and cell.style.comment:
            comment = escape_html(cell.style.comment)
            title_attr = f' title="{comment}"'
        if cell.formula:
            formula_escaped = escape_html(cell.formula)
            if title_attr:
                title_attr = f'{title_attr} | Formula: {formula_escaped}"'
            else:
                title_attr = f' title="Formula: {formula_escaped}"'

        data_attr = ""
        if cell.style and cell.style.number_format:
            number_format = escape_html(cell.style.number_format)
            data_attr = f' data-number-format="{number_format}"'
        tag = 'th' if is_header else 'td'
        return f'<{tag}{style_class}{span_attrs}{title_attr}{data_attr}{overflow_style}>{cell_content}</{tag}>'

    def _should_overflow_text(self, cell, row, col_idx):
        """
        检查单元格是否应该应用文字溢出显示（模拟Excel行为）。

        条件：
        1. 单元格有文字内容
        2. 文字长度超过一定阈值（比如10个字符）
        3. 右边的单元格为空
        4. 单元格没有设置文字换行
        """
        if not cell.value:
            return False

        # 检查文字长度 (中文字符按2倍计算)
        cell_text = str(cell.value).strip()
        # 计算显示宽度：中文字符按2倍计算，英文字符按1倍计算
        display_width = sum(2 if ord(c) > 127 else 1 for c in cell_text)
        if display_width <= 8:  # 短文字不需要溢出 (降低阈值以适应中文)
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
