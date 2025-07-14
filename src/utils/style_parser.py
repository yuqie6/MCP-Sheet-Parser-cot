"""
用于从 openpyxl 对象解析单元格样式的工具函数。
"""
from typing import Any
from openpyxl.cell.cell import Cell as OpenpyxlCell, MergedCell as OpenpyxlMergedCell
from openpyxl.styles.colors import Color
from src.models.table_model import Style, RichTextFragment, RichTextFragmentStyle, CellValue

# 定义一个类型提示，表示既可以是普通单元格也可以是合并单元格
CellLike = OpenpyxlCell | OpenpyxlMergedCell

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
    if cell.fill:
        background_color = extract_fill_color(cell.fill)
        if background_color:
            style.background_color = background_color
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

def get_border_style(border_side) -> str:
    """
    将 openpyxl 边框样式转换为 CSS 边框样式。
    """
    if not border_side or not border_side.style:
        return ""
    border_style_map = {
        'thin': '1px solid', 'medium': '2px solid', 'thick': '2px solid',
        'double': '3px double', 'dotted': '1px dotted', 'dashed': '1px dashed',
        'hair': '1px solid', 'mediumDashed': '2px dashed', 'dashDot': '1px dashed',
        'mediumDashDot': '2px dashed', 'dashDotDot': '1px dashed',
        'mediumDashDotDot': '2px dashed', 'slantDashDot': '1px dashed'
    }
    style_str = border_style_map.get(border_side.style)
    if not style_str:
        return ""
    color = extract_color(border_side.color) if border_side.color else None
    final_color = color if color else "#000000"
    return f"{style_str} {final_color}"

def extract_fill_color(fill) -> str | None:
    """
    针对多种填充类型增强的填充色提取。
    """
    if not fill:
        return None
    try:
        if hasattr(fill, 'patternType') and fill.patternType:
            if fill.patternType == 'solid' and hasattr(fill, 'start_color') and fill.start_color:
                color = extract_color(fill.start_color)
                if color and color != "#FFFFFF":  # 保留黑色背景，只过滤白色
                    return color
            elif fill.patternType in ['lightGray', 'mediumGray', 'darkGray']:
                return {'lightGray': "#F2F2F2", 'mediumGray': "#D9D9D9", 'darkGray': "#BFBFBF"}.get(fill.patternType)
            elif hasattr(fill, 'fgColor') and fill.fgColor:
                color = extract_color(fill.fgColor)
                if color and color != "#FFFFFF":  # 保留黑色背景，只过滤白色
                    return color
        if hasattr(fill, 'type') and fill.type == 'gradient':
            if hasattr(fill, 'stop') and fill.stop:
                stop = fill.stop[0] if isinstance(fill.stop, (list, tuple)) and len(fill.stop) > 0 else fill.stop
                if hasattr(stop, 'color') and not isinstance(stop, (list, tuple)):
                    color = extract_color(stop.color)
                    if color and color != "#FFFFFF":  # 保留黑色背景，只过滤白色
                        return color
    except Exception:
        pass
    return None

def extract_color(color_obj) -> str | None:
    """
    使用 openpyxl 的统一 value 接口简化颜色提取。
    """
    if not color_obj:
        return None
    try:
        value = color_obj.value
        if isinstance(value, str) and len(value) >= 6:
            return f"#{value[-6:]}"
        elif isinstance(value, int):
            # 检查是否为默认颜色索引
            if value == 1:  # 索引1通常是白色，但在Excel中默认文字应该是黑色
                # 检查是否为自动/默认颜色
                if hasattr(color_obj, 'auto') and color_obj.auto:
                    return None  # 返回None让系统使用默认黑色
                # 检查是否为主题颜色
                if hasattr(color_obj, 'theme') and color_obj.theme is not None:
                    # 主题颜色1通常是深色文字
                    if color_obj.theme == 1:
                        return "#000000"  # 返回黑色
            return get_color_by_index(value)
    except Exception:
        pass
    return None

def get_color_by_index(index: int) -> str:
    """
    标准色与主题色统一的索引颜色映射。
    """
    color_map = {
        0: "#000000", 1: "#FFFFFF", 2: "#FF0000", 3: "#00FF00", 4: "#0000FF",
        5: "#FFFF00", 6: "#FF00FF", 7: "#00FFFF", 8: "#000000", 9: "#FFFFFF",
        10: "#FF0000", 64: "#000000"
    }
    return color_map.get(index, "#000000")

def extract_number_format(cell) -> str:
    """
    增强的数字格式提取。
    """
    try:
        if cell.number_format and cell.number_format != 'General':
            format_str = cell.number_format
            format_mappings = {
            }
            return format_mappings.get(format_str, format_str) or ""
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
