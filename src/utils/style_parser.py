"""
Utility functions for parsing cell styles from openpyxl objects.
"""
from typing import Any
from src.models.table_model import Style

def extract_style(cell) -> Style:
    """
    Extracts comprehensive style information from an openpyxl cell.
    """
    style = Style()
    if cell.font:
        font = cell.font
        style.bold = font.bold if font.bold is not None else False
        style.italic = font.italic if font.italic is not None else False
        style.underline = font.underline is not None and font.underline != 'none'
        style.font_size = font.size if font.size else None
        style.font_name = font.name if font.name else None
        if font.color:
            font_color = extract_color(font.color)
            if font_color and font_color != "#000000":
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
        except:
            pass
    return style

def get_border_style(border_side) -> str:
    """
    Converts openpyxl border style to CSS border style.
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
    Enhanced fill color extraction for various fill types.
    """
    if not fill:
        return None
    try:
        if hasattr(fill, 'patternType') and fill.patternType:
            if fill.patternType == 'solid' and hasattr(fill, 'start_color') and fill.start_color:
                color = extract_color(fill.start_color)
                if color and color not in ["#FFFFFF", "#000000"]:
                    return color
            elif fill.patternType in ['lightGray', 'mediumGray', 'darkGray']:
                return {'lightGray': "#F2F2F2", 'mediumGray': "#D9D9D9", 'darkGray': "#BFBFBF"}.get(fill.patternType)
            elif hasattr(fill, 'fgColor') and fill.fgColor:
                color = extract_color(fill.fgColor)
                if color and color not in ["#FFFFFF", "#000000"]:
                    return color
        if hasattr(fill, 'type') and fill.type == 'gradient':
            if hasattr(fill, 'stop') and fill.stop:
                stop = fill.stop[0] if isinstance(fill.stop, (list, tuple)) and len(fill.stop) > 0 else fill.stop
                if hasattr(stop, 'color') and not isinstance(stop, (list, tuple)):
                    color = extract_color(stop.color)
                    if color and color not in ["#FFFFFF", "#000000"]:
                        return color
    except Exception:
        pass
    return None

def extract_color(color_obj) -> str | None:
    """
    Simplified color extraction using openpyxl's unified value interface.
    """
    if not color_obj:
        return None
    try:
        value = color_obj.value
        if isinstance(value, str) and len(value) >= 6:
            return f"#{value[-6:]}"
        elif isinstance(value, int):
            return get_color_by_index(value)
    except Exception:
        pass
    return None

def get_color_by_index(index: int) -> str:
    """
    Unified indexed color map for standard and theme colors.
    """
    color_map = {
        0: "#000000", 1: "#FFFFFF", 2: "#FF0000", 3: "#00FF00", 4: "#0000FF",
        5: "#FFFF00", 6: "#FF00FF", 7: "#00FFFF", 8: "#000000", 9: "#FFFFFF",
        10: "#FF0000", 64: "#000000"
    }
    return color_map.get(index, "#000000")

def extract_number_format(cell) -> str:
    """
    Enhanced number format extraction.
    """
    try:
        if cell.number_format and cell.number_format != 'General':
            format_str = cell.number_format
            format_mappings = {
            }
            return format_mappings.get(format_str, format_str) or ""
    except:
        pass
    return ""

def extract_hyperlink(cell) -> str | None:
    """
    Enhanced hyperlink extraction.
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
    Converts a Style object to a dictionary.
    """
    if not style:
        return {}
    
    style_dict = {}
    
    # Use vars() to dynamically get attributes, but filter out defaults to keep it clean.
    default_style = Style()
    for attr, value in vars(style).items():
        if value != getattr(default_style, attr):
             style_dict[attr] = value
             
    return style_dict
