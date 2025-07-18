"""
用于从 openpyxl 对象解析单元格样式的工具函数。
"""
from typing import Any, TypeAlias
from openpyxl.cell.cell import Cell as OpenpyxlCell, MergedCell as OpenpyxlMergedCell
from src.models.table_model import Style, RichTextFragment, RichTextFragmentStyle, CellValue
from src.utils.color_utils import extract_color, get_color_brightness, has_sufficient_contrast, get_color_by_index, get_theme_color, apply_tint, apply_smart_color_matching
from src.utils.border_utils import get_border_style

# 定义一个类型提示，表示既可以是普通单元格也可以是合并单元格
CellLike: TypeAlias = OpenpyxlCell | OpenpyxlMergedCell

def _extract_rich_text(cell: CellLike) -> list[RichTextFragment]:
    """从单元格中提取富文本片段。"""
    fragments = []
    # 合并单元格没有值，只有合并区域左上角的单元格有值。
    # 解析器应已处理此情况，这里做二次保护。
    if not hasattr(cell, 'value') or cell.value is None:
        return []

    if isinstance(cell.value, list):
        # 处理富文本
        for fragment in cell.value:
            font = getattr(fragment, 'font', None)
            if font:
                style = RichTextFragmentStyle(
                    bold=getattr(font, 'bold', False) or False,
                    italic=getattr(font, 'italic', False) or False,
                    underline=getattr(font, 'underline', None) is not None and getattr(font, 'underline', None) != 'none',
                    font_name=getattr(font, 'name', None),
                    font_size=getattr(font, 'size', None),
                    font_color=extract_color(getattr(font, 'color', None)) if getattr(font, 'color', None) else None
                )
            else:
                style = RichTextFragmentStyle()
            fragments.append(RichTextFragment(text=getattr(fragment, 'text', ''), style=style))
    else:
        # 处理纯文本
        font = getattr(cell, 'font', None)
        if font:
            style = RichTextFragmentStyle(
                bold=getattr(font, 'bold', False) or False,
                italic=getattr(font, 'italic', False) or False,
                underline=getattr(font, 'underline', None) is not None and getattr(font, 'underline', None) != 'none',
                font_name=getattr(font, 'name', None),
                font_size=getattr(font, 'size', None),
                font_color=extract_color(getattr(font, 'color', None)) if getattr(font, 'color', None) else None
            )
        else:
            style = RichTextFragmentStyle()
        fragments.append(RichTextFragment(text=str(cell.value), style=style))
    return fragments

def extract_cell_value(cell: CellLike) -> CellValue:
    """提取单元格的值，如有富文本则处理为富文本片段。"""
    if not hasattr(cell, 'value'):
        return None
    if isinstance(cell.value, list):
        return _extract_rich_text(cell)
    return cell.value

def extract_style(cell: CellLike) -> Style:
    """
    从 openpyxl 单元格中提取完整的样式信息。
    兼容普通单元格和合并单元格。
    """
    style = Style()
    if not hasattr(cell, 'has_style') or not cell.has_style:
        return style
    
    # 提取字体样式
    if cell.font:
        font = cell.font
        style.bold = font.bold if font.bold is not None else False
        style.italic = font.italic if font.italic is not None else False
        style.underline = font.underline is not None and font.underline != 'none'
        style.font_size = font.size if font.size else None
        style.font_name = font.name if font.name else None
        if font.color:
            font_color = extract_color(font.color)
            if font_color:
                style.font_color = font_color

    # 提取背景色
    if cell.fill:
        background_color = extract_fill_color(cell.fill)
        if background_color:
            style.background_color = background_color

    # 智能匹配字体色与背景色
    style = apply_smart_color_matching(style)
    if cell.alignment:
        alignment = cell.alignment
        if alignment.horizontal:
            style.text_align = alignment.horizontal
        if alignment.vertical:
            style.vertical_align = alignment.vertical
        style.wrap_text = alignment.wrap_text if alignment.wrap_text is not None else False
    if cell.border:
        border = cell.border
        style.border_top = get_border_style(border.top) if border.top else ""
        style.border_bottom = get_border_style(border.bottom) if border.bottom else ""
        style.border_left = get_border_style(border.left) if border.left else ""
        style.border_right = get_border_style(border.right) if border.right else ""
        border_color = None
        for border_side in [border.top, border.bottom, border.left, border.right]:
            if border_side and border_side.color:
                extracted_color = extract_color(border_side.color)
                if extracted_color:
                    border_color = extracted_color
                    break
        if border_color:
            style.border_color = border_color
        else:
            style.border_color = Style().border_color
    style.number_format = extract_number_format(cell)
    if cell.hyperlink:
        style.hyperlink = extract_hyperlink(cell)
    if cell.comment:
        try:
            if hasattr(cell.comment, 'text'):
                style.comment = str(cell.comment.text)
            elif hasattr(cell.comment, 'content'):
                style.comment = str(cell.comment.content)
        except (AttributeError, TypeError):
            pass
    return style

def extract_fill_color(fill) -> str | None:
    """
    针对多种填充类型增强的填充色提取。
    只有当填充类型为solid或其他有效类型时才提取背景色。
    """
    if not fill:
        return None
    try:
        # 检查填充类型，只有solid等有效类型才提取背景色
        if hasattr(fill, 'patternType') and fill.patternType:
            if fill.patternType == 'solid':
                # 优先检查start_color
                if hasattr(fill, 'start_color') and fill.start_color:
                    color = extract_color(fill.start_color)
                    if color:
                        return color
                # 备用检查fgColor
                elif hasattr(fill, 'fgColor') and fill.fgColor:
                    color = extract_color(fill.fgColor)
                    if color:
                        return color
            elif fill.patternType in ['lightGray', 'mediumGray', 'darkGray']:
                return {'lightGray': "#F2F2F2", 'mediumGray': "#D9D9D9", 'darkGray': "#BFBFBF"}.get(fill.patternType)

        # 处理渐变填充
        if hasattr(fill, 'type') and fill.type == 'gradient':
            if hasattr(fill, 'stop') and fill.stop:
                stop = fill.stop[0] if isinstance(fill.stop, (list, tuple)) and len(fill.stop) > 0 else fill.stop
                if hasattr(stop, 'color') and not isinstance(stop, (list, tuple)):
                    color = extract_color(stop.color)
                    if color:
                        return color

        # 重要：如果patternType为None，不提取背景色
        # 这避免了提取那些有颜色信息但不应该显示背景的单元格

    except Exception:
        pass
    return None



def extract_number_format(cell) -> str:
    """
    增强的数字格式提取。
    """
    try:
        if cell.number_format and cell.number_format != 'General':
            return cell.number_format
    except (AttributeError, ValueError):
        pass
    return ""

def extract_hyperlink(cell) -> str | None:
    """
    增强的超链接提取。
    """
    if not cell.hyperlink:
        return None
    try:
        if hasattr(cell.hyperlink, 'target') and cell.hyperlink.target:
            target = cell.hyperlink.target
            if isinstance(target, str) and (target.startswith(('http://', 'https://', 'ftp://', 'mailto:', 'file://')) or '.' in target):
                return target
        if hasattr(cell.hyperlink, 'location') and cell.hyperlink.location:
            location = cell.hyperlink.location
            if isinstance(location, str):
                return f"#{location}"
        if hasattr(cell.hyperlink, 'display') and cell.hyperlink.display:
            return str(cell.hyperlink.display)
    except Exception:
        pass
    return None

def style_to_dict(style: Style) -> dict[str, Any]:
    """
    将 Style 对象转换为字典。
    """
    if not style:
        return {}
    
    style_dict = {}
    
    # 使用 vars() 动态获取属性，但过滤默认值以保持简洁。
    default_style = Style()
    for attr, value in vars(style).items():
        if value != getattr(default_style, attr):
             style_dict[attr] = value
             
    return style_dict
