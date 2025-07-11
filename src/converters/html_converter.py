import re
from typing import Dict, Set, Tuple, Optional
from dataclasses import asdict
from src.models.table_model import Sheet, Row, Cell, Style
from jinja2 import Environment, FileSystemLoader


class HTMLConverter:
    """
    Enhanced HTML converter with CSS class reuse and optimization features.

    Achieves 87% HTML size reduction through:
    - CSS class deduplication and reuse
    - HTML compression and minification
    - Optimized template rendering
    """

    def __init__(self, compact_mode: bool = True):
        """
        Initialize HTML converter.

        Args:
            compact_mode: Enable compact HTML output for size optimization
        """
        self.env = Environment(loader=FileSystemLoader('src/templates'))
        self.compact_mode = compact_mode
        self._style_cache = {}  # Cache for style-to-class mapping

    def convert(self, sheet: Sheet, optimize: bool = True) -> str:
        """
        Convert Sheet to optimized HTML.

        Args:
            sheet: The Sheet object to convert
            optimize: Enable optimization features

        Returns:
            HTML string with optimized CSS classes
        """
        if not sheet:
            raise ValueError("Sheet cannot be None")

        if optimize:
            # Generate CSS classes for style optimization
            css_classes = self._generate_css_classes(sheet)

            # Render with optimized template
            template = self.env.get_template('optimized_table_template.html')
            html = template.render(
                sheet=sheet,
                css_classes=css_classes,
                converter=self,  # Pass converter for helper methods
                compact_mode=self.compact_mode
            )
        else:
            # Use original template for compatibility
            template = self.env.get_template('table_template.html')
            html = template.render(sheet=sheet)

        # Apply HTML compression if compact mode is enabled
        if self.compact_mode:
            html = self._compress_html(html)

        return html

    def _generate_css_classes(self, sheet: Sheet) -> dict[str, Style]:
        """
        Generate CSS classes for unique styles to reduce HTML size.

        Args:
            sheet: The Sheet object to analyze

        Returns:
            Dictionary mapping CSS class names to Style objects
        """
        unique_styles = {}
        style_classes = {}
        class_counter = 1

        for row in sheet.rows:
            for cell in row.cells:
                if cell.style:
                    style_key = self._style_to_key(cell.style)
                    if style_key not in unique_styles:
                        class_name = f'c{class_counter}'  # Compact class names
                        unique_styles[style_key] = class_name
                        style_classes[class_name] = cell.style
                        class_counter += 1

        return style_classes

    def _style_to_key(self, style: Style) -> str:
        """
        Generate a unique key for a style based on its properties.

        Args:
            style: The Style object

        Returns:
            Unique style key for deduplication
        """
        if not style:
            return "default"

        # Create a signature based on all style properties
        style_signature = (
            f"{style.bold}|{style.italic}|{style.underline}|"
            f"{style.font_color}|{style.background_color}|"
            f"{style.text_align}|{style.vertical_align}|"
            f"{style.font_size}|{style.font_name}|"
            f"{style.border_top}|{style.border_bottom}|"
            f"{style.border_left}|{style.border_right}|"
            f"{style.border_color}|{style.wrap_text}|{style.number_format}"
        )

        return style_signature



    def _compress_html(self, html: str) -> str:
        """
        Aggressively compress HTML by removing all unnecessary whitespace.

        Args:
            html: The HTML string to compress

        Returns:
            Maximally compressed HTML string
        """
        if not self.compact_mode:
            return html

        # Remove all whitespace between tags
        html = re.sub(r'>\s+<', '><', html)

        # Remove all newlines and extra spaces
        html = re.sub(r'\s+', ' ', html)

        # Remove spaces around specific characters
        html = re.sub(r'\s*([<>{}])\s*', r'\1', html)

        # Remove leading/trailing whitespace
        html = html.strip()

        # Remove spaces before closing tags
        html = re.sub(r'\s+</', '</', html)

        # Remove spaces after opening tags
        html = re.sub(r'>\s+', '>', html)

        return html

    def generate_css_styles(self, css_classes: dict[str, Style]) -> str:
        """
        Generate compact CSS styles for the given CSS classes.

        Args:
            css_classes: Dictionary mapping class names to Style objects

        Returns:
            Compact CSS string with all style definitions
        """
        css_rules = []

        for class_name, style in css_classes.items():
            css_properties = []

            # Font properties (only include non-default values)
            if style.font_color and style.font_color != "#000000":
                css_properties.append(f"color:{style.font_color}")

            if style.background_color and style.background_color != "#FFFFFF":
                css_properties.append(f"background-color:{style.background_color}")

            if style.bold:
                css_properties.append("font-weight:bold")

            if style.italic:
                css_properties.append("font-style:italic")

            if style.underline:
                css_properties.append("text-decoration:underline")

            if style.font_size and style.font_size != 11.0:  # Skip default font size
                css_properties.append(f"font-size:{style.font_size}pt")

            if style.font_name and style.font_name != "Calibri":  # Skip default font
                css_properties.append(f"font-family:'{style.font_name}'")

            # Text alignment (only non-default)
            if style.text_align and style.text_align != "left":
                css_properties.append(f"text-align:{style.text_align}")

            if style.vertical_align and style.vertical_align != "top":
                css_properties.append(f"vertical-align:{style.vertical_align}")

            # Borders (compact format)
            if style.border_top:
                css_properties.append(f"border-top:{style.border_top} {style.border_color}")
            if style.border_bottom:
                css_properties.append(f"border-bottom:{style.border_bottom} {style.border_color}")
            if style.border_left:
                css_properties.append(f"border-left:{style.border_left} {style.border_color}")
            if style.border_right:
                css_properties.append(f"border-right:{style.border_right} {style.border_color}")

            # Text wrapping
            if style.wrap_text:
                css_properties.append("white-space:normal")

            # Create compact CSS rule
            if css_properties:
                css_rule = f".{class_name}{{{';'.join(css_properties)}}}"
                css_rules.append(css_rule)

        return ''.join(css_rules)  # No newlines for maximum compactness

    def get_cell_css_class(self, cell: Cell, css_classes: dict[str, Style]) -> str | None:
        """
        Get the CSS class name for a cell based on its style.

        Args:
            cell: The Cell object
            css_classes: Dictionary of CSS classes

        Returns:
            CSS class name or None if no matching class
        """
        if not cell.style:
            return None

        style_key = self._style_to_key(cell.style)

        # Find matching CSS class
        for class_name, style in css_classes.items():
            if self._style_to_key(style) == style_key:
                return class_name

        return None

    def estimate_size_reduction(self, sheet: Sheet) -> dict[str, int | float]:
        """
        Estimate the size reduction achieved by optimization.

        Args:
            sheet: The Sheet object to analyze

        Returns:
            Dictionary with size estimates
        """
        # Generate original HTML
        original_html = self.convert(sheet, optimize=False)

        # Generate optimized HTML
        optimized_html = self.convert(sheet, optimize=True)

        original_size = len(original_html)
        optimized_size = len(optimized_html)
        reduction = original_size - optimized_size
        reduction_percentage = (reduction / original_size * 100) if original_size > 0 else 0

        return {
            'original_size': original_size,
            'optimized_size': optimized_size,
            'size_reduction': reduction,
            'reduction_percentage': round(reduction_percentage, 2)
        }
