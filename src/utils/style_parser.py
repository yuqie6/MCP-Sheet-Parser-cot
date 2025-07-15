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
    
    # 应用智能边框颜色处理，确保边框一致性
    if color:
        # 转换白色和接近白色的边框颜色
        if color.upper() in ['#FFFFFF', '#FFF', 'WHITE', '#FEFEFE', '#FDFDFD']:
            final_color = '#E0E0E0'
        # 转换与旧默认边框太接近的颜色
        elif color.upper() in ['#D8D8D8', '#DADADA', '#DBDBDB']:
            final_color = '#E0E0E0'
        else:
            final_color = color
    else:
        final_color = "#E0E0E0"  # 与默认表格边框保持一致
        
    return f"{style_str} {final_color}"

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

def extract_color(color_obj) -> str | None:
    """
    增强的颜色提取，支持主题颜色、tint和RGB。
    """
    if not color_obj:
        return None

    try:
        # 1. 处理RGB颜色
        if hasattr(color_obj, 'rgb') and color_obj.rgb:
            rgb_value = color_obj.rgb
            if isinstance(rgb_value, str) and len(rgb_value) >= 6:
                # 处理ARGB格式（如FFFF0000）- 移除前两位Alpha通道
                if len(rgb_value) == 8 and rgb_value.startswith('FF'):
                    return f"#{rgb_value[2:]}"
                # 标准RGB格式
                elif len(rgb_value) == 6:
                    return f"#{rgb_value}"
                # 其他情况取最后6位
                else:
                    return f"#{rgb_value[-6:]}"

        # 2. 处理主题颜色
        if hasattr(color_obj, 'theme') and color_obj.theme is not None:
            theme_color = get_theme_color(color_obj.theme)
            if theme_color:
                # 应用tint调整
                if hasattr(color_obj, 'tint') and color_obj.tint != 0:
                    return apply_tint(theme_color, color_obj.tint)
                return theme_color

        # 3. 处理索引颜色
        if hasattr(color_obj, 'indexed') and color_obj.indexed is not None:
            return get_color_by_index(color_obj.indexed)

        # 4. 处理value属性（向后兼容）
        if hasattr(color_obj, 'value') and color_obj.value is not None:
            value = color_obj.value
            if isinstance(value, str) and len(value) >= 6:
                return f"#{value[-6:]}"
            elif isinstance(value, int):
                return get_color_by_index(value)

        # 5. 检查auto属性
        if hasattr(color_obj, 'auto') and color_obj.auto:
            return None  # 自动颜色，让系统决定

    except Exception:
        pass

    return None

def get_theme_color(theme_index: int) -> str | None:
    """
    Excel主题颜色映射。
    基于实际Excel文件的主题颜色方案。
    """
    theme_colors = {
        0: "#FFFFFF",  # 背景1（白色）
        1: "#000000",  # 文本1（黑色）
        2: "#E7E6E6",  # 背景2（浅灰）
        3: "#44546A",  # 文本2（深蓝灰）
        4: "#5B9BD5",  # 强调1（蓝色）
        5: "#70AD47",  # 强调2（绿色）
        6: "#A5A5A5",  # 强调3（灰色）
        7: "#FFC000",  # 强调4（橙色）
        8: "#4472C4",  # 强调5（深蓝）
        9: "#70AD47",  # 强调6（绿色）- 根据截图修正为绿色
        10: "#FF0000", # 超链接（红色）
        11: "#800080"  # 已访问超链接（紫色）
    }
    return theme_colors.get(theme_index)

def apply_tint(base_color: str, tint: float) -> str:
    """
    对基础颜色应用tint调整。
    tint > 0: 变亮（向白色混合）
    tint < 0: 变暗（向黑色混合）
    """
    if not base_color or not base_color.startswith('#'):
        return base_color

    try:
        # 解析RGB值
        r = int(base_color[1:3], 16)
        g = int(base_color[3:5], 16)
        b = int(base_color[5:7], 16)

        # 应用tint调整
        if tint > 0:
            # 向白色混合
            r = int(r + (255 - r) * tint)
            g = int(g + (255 - g) * tint)
            b = int(b + (255 - b) * tint)
        elif tint < 0:
            # 向黑色混合
            r = int(r * (1 + tint))
            g = int(g * (1 + tint))
            b = int(b * (1 + tint))

        # 确保值在0-255范围内
        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))

        return f"#{r:02X}{g:02X}{b:02X}"

    except (ValueError, IndexError):
        return base_color

def apply_smart_color_matching(style: Style) -> Style:
    """
    智能匹配字体色与背景色，确保对比度和可读性。
    只在真正需要的时候才应用智能匹配。
    """
    # 如果没有背景色，不需要调整
    if not style.background_color:
        return style

    # 如果已经有字体色，检查对比度
    if style.font_color:
        # 如果对比度足够，保持原样
        if has_sufficient_contrast(style.font_color, style.background_color):
            return style
        # 对比度不足时才进行调整
    else:
        # 没有字体色时，只有在背景色很深的情况下才设置白色字体
        bg_brightness = get_color_brightness(style.background_color)
        if bg_brightness < 64:  # 只有非常深的背景才自动设置白色字体
            style.font_color = "#FFFFFF"
        return style

    # 根据背景色的亮度决定字体色（只在对比度不足时）
    bg_brightness = get_color_brightness(style.background_color)

    if bg_brightness < 128:  # 深色背景
        style.font_color = "#FFFFFF"  # 使用白色字体
    else:  # 浅色背景
        style.font_color = "#000000"  # 使用黑色字体

    return style

def get_color_brightness(color: str) -> int:
    """
    计算颜色的亮度（0-255）。
    使用感知亮度公式：0.299*R + 0.587*G + 0.114*B
    """
    if not color or not color.startswith('#'):
        return 128  # 默认中等亮度

    try:
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)
        return int(0.299 * r + 0.587 * g + 0.114 * b)
    except (ValueError, IndexError):
        return 128

def has_sufficient_contrast(color1: str, color2: str) -> bool:
    """
    检查两个颜色之间是否有足够的对比度。
    """
    brightness1 = get_color_brightness(color1)
    brightness2 = get_color_brightness(color2)
    return abs(brightness1 - brightness2) > 100  # 对比度阈值

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
