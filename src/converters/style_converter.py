from typing import Optional
from src.models.table_model import Style


class StyleConverter:
    """Converts Style objects to CSS style strings."""
    
    def style_to_css(self, style: Optional[Style]) -> str:
        """
        Convert a Style object to a CSS style string.
        
        Args:
            style: Style object containing formatting information
            
        Returns:
            CSS style string ready for use in HTML style attribute
        """
        if not style:
            return ""
        
        css_parts = []
        
        # Font properties
        if style.font_name:
            css_parts.append(f"font-family: '{style.font_name}'")
        
        if style.font_size:
            css_parts.append(f"font-size: {style.font_size}pt")
        
        if style.font_color:
            css_parts.append(f"color: {style.font_color}")
        
        # Text formatting
        if style.bold:
            css_parts.append("font-weight: bold")
        
        if style.italic:
            css_parts.append("font-style: italic")
        
        if style.underline:
            css_parts.append("text-decoration: underline")
        elif style.text_decoration:
            css_parts.append(f"text-decoration: {style.text_decoration}")
        
        # Background color
        if style.fill_color:
            css_parts.append(f"background-color: {style.fill_color}")
        
        # Text alignment
        if style.alignment:
            css_parts.append(f"text-align: {style.alignment}")
        
        if style.vertical_alignment:
            vertical_map = {
                'top': 'top',
                'middle': 'middle', 
                'bottom': 'bottom'
            }
            if style.vertical_alignment in vertical_map:
                css_parts.append(f"vertical-align: {vertical_map[style.vertical_alignment]}")
        
        # Border properties
        border_css = self._build_border_css(style)
        if border_css:
            css_parts.append(border_css)
        
        # Cell dimensions
        if style.width:
            css_parts.append(f"width: {style.width}")
        
        if style.height:
            css_parts.append(f"height: {style.height}")
        
        if style.padding:
            css_parts.append(f"padding: {style.padding}")
        
        return "; ".join(css_parts)
    
    def _build_border_css(self, style: Style) -> str:
        """Build border CSS from style properties."""
        border_parts = []
        
        if style.border_style:
            border_parts.append(style.border_style)
        else:
            border_parts.append("solid")  # Default border style
        
        if style.border_width:
            border_parts.append(style.border_width)
        else:
            border_parts.append("1px")  # Default border width
        
        if style.border_color:
            border_parts.append(style.border_color)
        else:
            border_parts.append("#000000")  # Default border color
        
        # Only return border CSS if at least one border property is set
        if style.border_style or style.border_width or style.border_color:
            return f"border: {' '.join(border_parts)}"
        
        return ""
    
    def get_table_css(self) -> str:
        """
        Get default CSS styles for the table structure.
        
        Returns:
            CSS string with default table styling
        """
        return """
        table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
        }
        
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
            vertical-align: top;
        }
        
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        
        .merged-cell {
            background-color: #f9f9f9;
        }
        """
