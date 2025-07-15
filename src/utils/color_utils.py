"""
颜色处理工具模块

提供统一的颜色处理功能，避免代码重复。
"""

import re
from typing import Optional, Dict

# Excel主题颜色映射（Office 2016+ 默认主题）
EXCEL_THEME_COLORS = {
    'accent1': '#5B9BD5',  # 蓝色
    'accent2': '#70AD47',  # 绿色  
    'accent3': '#FFC000',  # 黄色/橙色
    'accent4': '#E15759',  # 红色
    'accent5': '#4472C4',  # 深蓝色
    'accent6': '#70AD47',  # 绿色（在某些Excel文件中）
    'dk1': '#000000',      # 深色1（黑色）
    'lt1': '#FFFFFF',      # 浅色1（白色）
    'dk2': '#44546A',      # 深色2（深灰蓝）
    'lt2': '#E7E6E6',      # 浅色2（浅灰）
    'bg1': '#FFFFFF',      # 背景1
    'bg2': '#E7E6E6',      # 背景2
    'tx1': '#000000',      # 文本1
    'tx2': '#44546A',      # 文本2
    'hlink': '#0563C1',    # 超链接
    'folHlink': '#954F72', # 已访问超链接
}

# 图表系列的默认颜色顺序
DEFAULT_CHART_COLORS = [
    EXCEL_THEME_COLORS['accent2'],  # 绿色
    EXCEL_THEME_COLORS['accent1'],  # 蓝色
    EXCEL_THEME_COLORS['accent3'],  # 橙色
    EXCEL_THEME_COLORS['accent4'],  # 红色
    EXCEL_THEME_COLORS['accent5'],  # 深蓝色
    '#FF6B35',  # 橙红色
]

# 颜色名称映射
COLOR_NAMES = {
    'BLACK': '#000000',
    'WHITE': '#FFFFFF',
    'RED': '#FF0000',
    'GREEN': '#00FF00',
    'BLUE': '#0000FF',
    'YELLOW': '#FFFF00',
    'CYAN': '#00FFFF',
    'MAGENTA': '#FF00FF',
    'GRAY': '#808080',
    'GREY': '#808080',
    'SILVER': '#C0C0C0',
    'MAROON': '#800000',
    'OLIVE': '#808000',
    'LIME': '#00FF00',
    'AQUA': '#00FFFF',
    'TEAL': '#008080',
    'NAVY': '#000080',
    'FUCHSIA': '#FF00FF',
    'PURPLE': '#800080',
}


def normalize_color(color: str) -> str:
    """
    标准化颜色格式为#RRGGBB格式。
    
    参数：
        color: 颜色值（可能包含#前缀或不包含）
        
    返回：
        标准化的#RRGGBB格式颜色
    """
    if not color:
        return '#000000'
    
    # 移除前缀（如果有）
    clean_color = color.replace('#', '').replace('00', '', 1) if color.startswith('00') else color.replace('#', '')
    
    # 确保是6位十六进制
    if len(clean_color) == 6:
        return f'#{clean_color.upper()}'
    elif len(clean_color) == 8 and clean_color.startswith('00'):
        return f'#{clean_color[2:].upper()}'
    else:
        return '#000000'  # 默认黑色


def format_color(color: str, is_font_color: bool = False, is_border_color: bool = False) -> Optional[str]:
    """
    格式化颜色字符串。
    
    参数：
        color: 原始颜色字符串
        is_font_color: 是否为字体颜色
        is_border_color: 是否为边框颜色
        
    返回：
        格式化后的颜色字符串，失败时返回None
    """
    if not color:
        return '#000000' if is_font_color else None
    
    color = color.strip().upper()
    
    # 检查各种颜色格式
    if re.match(r'^#[0-9A-F]{6}$', color):
        formatted_color = color
    elif re.match(r'^#[0-9A-F]{3}$', color):
        formatted_color = f"#{color[1]}{color[1]}{color[2]}{color[2]}{color[3]}{color[3]}"
    elif re.match(r'^[0-9A-F]{6}$', color):
        formatted_color = f"#{color}"
    elif re.match(r'^[0-9A-F]{3}$', color):
        formatted_color = f"#{color[0]}{color[0]}{color[1]}{color[1]}{color[2]}{color[2]}"
    elif color in COLOR_NAMES:
        formatted_color = COLOR_NAMES[color]
    else:
        return '#000000' if is_font_color else None

    # 边框颜色特殊处理
    if is_border_color:
        if formatted_color.upper() in ['#FFFFFF', '#FFF', 'WHITE', '#FEFEFE', '#FDFDFD']:
            return '#E0E0E0'
        elif formatted_color.upper() in ['#D8D8D8', '#DADADA', '#DBDBDB']:
            return '#E0E0E0'
    
    return formatted_color


def convert_scheme_color_to_hex(scheme_color: str) -> str:
    """
    将Excel主题颜色转换为十六进制颜色。
    
    参数：
        scheme_color: Excel主题颜色名称
        
    返回：
        十六进制颜色字符串
    """
    return EXCEL_THEME_COLORS.get(scheme_color, '#70AD47')  # 默认绿色


def generate_pie_color_variants(base_color: str, count: int) -> list[str]:
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
        return (DEFAULT_CHART_COLORS * ((count // len(DEFAULT_CHART_COLORS)) + 1))[:count]