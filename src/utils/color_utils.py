"""
颜色处理工具模块

提供统一的颜色处理功能，避免代码重复。
"""

import re
from typing import Optional, Dict, TYPE_CHECKING
from functools import lru_cache

if TYPE_CHECKING:
    from src.models.table_model import Style

# Excel主题颜色映射（Office 2016+ 默认主题）
EXCEL_THEME_COLORS = {
    'accent1': '#4F81BD',
    'accent2': '#C0504D',
    'accent3': '#9BBB59',
    'accent4': '#8064A2',
    'accent5': '#4BACC6',
    'accent6': '#F79646',
    'dk1': '#000000',
    'lt1': '#FFFFFF',
    'dk2': '#1F497D',
    'lt2': '#EEECE1',
    'hlink': '#0000FF',
    'folHlink': '#800080',
    # bg1, bg2, tx1, tx2 are not in the standard theme file, but are common.
    # We will keep them for compatibility, using the most logical mappings.
    'bg1': '#FFFFFF',      # Corresponds to lt1
    'bg2': '#EEECE1',      # Corresponds to lt2
    'tx1': '#000000',      # Corresponds to dk1
    'tx2': '#1F497D',      # Corresponds to dk2
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


@lru_cache(maxsize=256)
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
    clean_color = color.lstrip('#')
    
    # 处理以'00'开头的8位颜色
    if len(clean_color) == 8 and clean_color.startswith('00'):
        clean_color = clean_color[2:]
    
    # 确保是6位十六进制
    if len(clean_color) == 6:
        return f'#{clean_color.upper()}'
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


@lru_cache(maxsize=128)
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


def generate_distinct_colors(count: int, existing_colors: Optional[list] = None) -> list[str]:
    """
    生成与现有颜色不同的新颜色。统一的颜色生成工具。
    
    参数：
        count: 需要生成的颜色数量
        existing_colors: 已存在的颜色列表
        
    返回：
        新颜色列表
    """
    if existing_colors is None:
        existing_colors = []
    
    # 预定义的颜色池，确保视觉上有足够的区分度
    color_pool = [
        '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
        '#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9',
        '#F8C471', '#82E0AA', '#F1948A', '#85C1E9', '#D7BDE2',
        '#A3E4D7', '#F9E79F', '#D5A6BD', '#AED6F1', '#A9DFBF',
        '#FAD7A0', '#E8DAEF', '#D1F2EB', '#FCF3CF', '#FADBD8'
    ]
    
    # 过滤掉已存在的颜色
    available_colors = [c for c in color_pool if c not in existing_colors]
    
    # 如果可用颜色不够，生成更多颜色
    if len(available_colors) < count:
        # 使用HSV色彩空间生成更多颜色
        import colorsys
        additional_needed = count - len(available_colors)
        for i in range(additional_needed):
            # 在HSV空间中均匀分布色相
            hue = (i * 137.5) % 360 / 360  # 使用黄金角度分布
            saturation = 0.7 + (i % 3) * 0.1  # 0.7, 0.8, 0.9
            value = 0.8 + (i % 2) * 0.1  # 0.8, 0.9
            
            rgb = colorsys.hsv_to_rgb(hue, saturation, value)
            hex_color = f"#{int(rgb[0]*255):02X}{int(rgb[1]*255):02X}{int(rgb[2]*255):02X}"
            
            if hex_color not in existing_colors and hex_color not in available_colors:
                available_colors.append(hex_color)
    
    return available_colors[:count]


def ensure_distinct_colors(colors: list[str], required_count: int) -> list[str]:
    """
    确保颜色列表有足够的不重复颜色。
    
    参数：
        colors: 原始颜色列表
        required_count: 需要的颜色数量
        
    返回：
        包含足够不重复颜色的列表
    """
    if not colors:
        return generate_distinct_colors(required_count)
    
    # 去重但保持顺序
    unique_colors = list(dict.fromkeys(colors))
    
    if len(unique_colors) >= required_count:
        return unique_colors[:required_count]
    
    # 生成额外的颜色
    additional_colors = generate_distinct_colors(
        required_count - len(unique_colors), 
        unique_colors
    )
    
    return unique_colors + additional_colors


# ===== 从 style_parser.py 迁移的颜色处理函数 =====

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


def apply_smart_color_matching(style: "Style") -> "Style":
    """
    智能匹配字体色与背景色，确保对比度和可读性。
    只在真正需要的时候才应用智能匹配。
    
    参数：
        style: Style对象，需要有background_color和font_color属性
        
    返回：
        调整后的Style对象
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