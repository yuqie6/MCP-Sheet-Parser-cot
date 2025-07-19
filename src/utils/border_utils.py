"""
边框样式处理工具模块

提供统一的边框样式处理功能，避免代码重复。
"""

import re
from src.utils.color_utils import extract_color

# 统一的边框样式映射
BORDER_STYLE_MAP = {
    "thin": ("1px", "solid"),
    "medium": ("2px", "solid"),
    "thick": ("3px", "solid"),
    "solid": ("1px", "solid"),
    "dashed": ("1px", "dashed"),
    "dotted": ("1px", "dotted"),
    "double": ("3px", "double"),
    "groove": ("2px", "groove"),
    "ridge": ("2px", "ridge"),
    "inset": ("2px", "inset"),
    "outset": ("2px", "outset"),
    "hair": ("1px", "solid"),
    "mediumdashed": ("2px", "dashed"),
    "dashdot": ("1px", "dashed"),
    "mediumdashdot": ("2px", "dashed"),
    "dashdotdot": ("1px", "dashed"),
    "mediumdashdotdot": ("2px", "dashed"),
    "slantdashdot": ("1px", "dashed"),
    # Excel 内部样式值映射
    1: "solid",
    2: "dashed",
    3: "dotted",
    4: "double",
    5: "solid",  # 粗线
    6: "solid",  # 中等线
}


def get_border_style(border_side) -> str:
    """
    将 openpyxl 边框样式转换为 CSS 边框样式。
    
    参数：
        border_side: openpyxl 边框对象
        
    返回：
        CSS 边框样式字符串
    """
    if not border_side or not border_side.style:
        return ""
    
    style_str = BORDER_STYLE_MAP.get(border_side.style)
    if not style_str:
        return ""
    
    # 如果是元组，说明是新的格式 (width, style)
    if isinstance(style_str, tuple):
        width, style_type = style_str
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
            
        return f"{width} {style_type} {final_color}"
    else:
        # 旧格式，直接返回样式名称
        return style_str


def parse_border_style_complete(border_style: str | None, border_color: str = "#E0E0E0") -> str:
    """
    解析完整的边框样式字符串。
    
    参数：
        border_style: 边框样式字符串
        border_color: 边框颜色
        
    返回：
        完整的CSS边框样式
    """
    if not border_style:
        return f"1px solid {border_color}"
    
    border_lower = border_style.lower()
    if border_lower in BORDER_STYLE_MAP:
        style_info = BORDER_STYLE_MAP[border_lower]
        if isinstance(style_info, tuple):
            width, style_type = style_info
            return f"{width} {style_type} {border_color}"
        else:
            return f"1px {style_info} {border_color}"
    
    # 解析自定义样式
    pattern = r'(\d+(?:\.\d+)?)(px|pt|em|rem)?\s*(solid|dashed|dotted|double|groove|ridge|inset|outset)?'
    match = re.search(pattern, border_style.lower())
    if match:
        width = match.group(1)
        unit = match.group(2) or "px"
        style_type = match.group(3) or "solid"
        return f"{width}{unit} {style_type} {border_color}"
    
    return f"1px solid {border_color}"


def get_xls_border_style_name(style_code: int) -> str:
    """
    将 XLS 边框样式代码转换为样式名称。
    
    参数：
        style_code: XLS 边框样式代码
        
    返回：
        边框样式名称
    """
    return BORDER_STYLE_MAP.get(style_code, "solid")


def format_border_color(color: str | None) -> str:
    """
    格式化边框颜色，确保边框一致性。
    
    参数：
        color: 原始颜色字符串
        
    返回：
        格式化后的边框颜色
    """
    if not color:
        return "#E0E0E0"
    
    # 转换白色和接近白色的边框颜色
    if color.upper() in ['#FFFFFF', '#FFF', 'WHITE', '#FEFEFE', '#FDFDFD']:
        return '#E0E0E0'
    # 转换与旧默认边框太接近的颜色
    elif color.upper() in ['#D8D8D8', '#DADADA', '#DBDBDB']:
        return '#E0E0E0'
    else:
        return color