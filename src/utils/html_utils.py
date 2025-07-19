"""
HTML工具模块

提供统一的HTML处理功能，避免代码重复。
"""



def escape_html(text: str) -> str:
    """
    转义HTML特殊字符。
    
    参数：
        text: 要转义的文本
        
    返回：
        转义后的文本
    """
    if not isinstance(text, str):
        text = str(text)
    return (text.replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&#x27;'))


def generate_style_attribute(css_parts: list[str]) -> str:
    """
    生成HTML样式属性。
    
    参数：
        css_parts: CSS样式部分列表
        
    返回：
        完整的style属性字符串
    """
    if not css_parts:
        return ""
    return f'style="{" ".join(css_parts)}"'


def generate_class_attribute(css_classes: list[str]) -> str:
    """
    生成HTML类属性。
    
    参数：
        css_classes: CSS类名列表
        
    返回：
        完整的class属性字符串
    """
    if not css_classes:
        return ""
    return f'class="{" ".join(css_classes)}"'


def create_html_element(tag: str, content: str = "", attributes: dict[str, str] | None = None, 
                       css_classes: list[str] | None = None, inline_styles: dict[str, str] | None = None) -> str:
    """
    创建HTML元素。
    
    参数：
        tag: HTML标签名
        content: 元素内容
        attributes: HTML属性字典
        css_classes: CSS类名列表
        inline_styles: 内联样式字典
        
    返回：
        完整的HTML元素字符串
    """
    attr_parts = []
    
    # 添加基本属性
    if attributes:
        for key, value in attributes.items():
            attr_parts.append(f'{key}="{escape_html(value)}"')
    
    # 添加CSS类
    if css_classes:
        attr_parts.append(generate_class_attribute(css_classes))
    
    # 添加内联样式
    if inline_styles:
        css_parts = [f"{key}: {value};" for key, value in inline_styles.items()]
        attr_parts.append(generate_style_attribute(css_parts))
    
    attr_str = " " + " ".join(attr_parts) if attr_parts else ""
    
    if content:
        return f'<{tag}{attr_str}>{content}</{tag}>'
    else:
        return f'<{tag}{attr_str}>'


def create_table_cell(content: str, is_header: bool = False, rowspan: int = 1, 
                     colspan: int = 1, css_classes: list[str] | None = None, 
                     inline_styles: dict[str, str] | None = None, title: str | None = None) -> str:
    """
    创建表格单元格。
    
    参数：
        content: 单元格内容
        is_header: 是否为表头单元格
        rowspan: 行跨度
        colspan: 列跨度
        css_classes: CSS类名列表
        inline_styles: 内联样式字典
        title: 工具提示文本
        
    返回：
        完整的表格单元格HTML
    """
    tag = 'th' if is_header else 'td'
    attributes = {}
    
    if rowspan > 1:
        attributes['rowspan'] = str(rowspan)
    if colspan > 1:
        attributes['colspan'] = str(colspan)
    if title:
        attributes['title'] = title
    
    return create_html_element(tag, content, attributes, css_classes, inline_styles)


def create_svg_element(width: int, height: int, content: str = "", 
                      css_classes: list[str] | None = None) -> str:
    """
    创建SVG元素。
    
    参数：
        width: SVG宽度
        height: SVG高度
        content: SVG内容
        css_classes: CSS类名列表
        
    返回：
        完整的SVG元素HTML
    """
    attributes = {
        'width': f'{width}px',
        'height': f'{height}px',
        'viewBox': f'0 0 {width} {height}',
        'xmlns': 'http://www.w3.org/2000/svg'
    }
    
    return create_html_element('svg', content, attributes, css_classes)


def compact_html(html: str) -> str:
    """
    压缩HTML内容。
    
    参数：
        html: 原始HTML
        
    返回：
        压缩后的HTML
    """
    lines = html.split('\n')
    compact_lines = [line.strip() for line in lines if line.strip()]
    return '\n'.join(compact_lines)