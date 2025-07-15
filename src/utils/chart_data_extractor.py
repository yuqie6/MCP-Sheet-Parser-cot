"""
图表数据提取工具模块

提供统一的图表数据提取功能，避免代码重复。
"""

from typing import Any, Dict, List, Optional
from src.utils.color_utils import convert_scheme_color_to_hex, generate_pie_color_variants


class ChartDataExtractor:
    """图表数据提取器，统一处理各种图表数据提取逻辑。"""
    
    def extract_series_y_data(self, series) -> List[float]:
        """
        提取系列的Y轴数据。
        
        参数：
            series: openpyxl图表系列对象
            
        返回：
            Y轴数据列表
        """
        y_data = []
        
        # 方法1：从series.val获取数值数据
        if hasattr(series, 'val') and series.val:
            y_data = self._extract_from_val(series.val)
        
        # 方法2：从series.yVal获取（备用方法）
        if not y_data and hasattr(series, 'yVal') and series.yVal:
            y_data = self._extract_from_val(series.yVal)
        
        return y_data
    
    def extract_series_x_data(self, series) -> List[str]:
        """
        提取系列的X轴数据。
        
        参数：
            series: openpyxl图表系列对象
            
        返回：
            X轴数据列表
        """
        x_data = []
        
        # 方法1：从series.cat获取分类数据
        if hasattr(series, 'cat') and series.cat:
            x_data = self._extract_from_cat(series.cat)
        
        # 方法2：从series.xVal获取（备用方法）
        if not x_data and hasattr(series, 'xVal') and series.xVal:
            x_data = self._extract_from_xval(series.xVal)
        
        return x_data
    
    def extract_series_color(self, series) -> Optional[str]:
        """
        提取系列的颜色信息。
        
        参数：
            series: openpyxl图表系列对象
            
        返回：
            颜色的十六进制字符串，提取失败时返回None
        """
        try:
            # 尝试从图形属性中获取颜色
            color = self._extract_from_graphics_properties(series)
            if color:
                return color
            
            # 尝试从spPr属性获取
            color = self._extract_from_sp_pr(series)
            if color:
                return color
            
        except Exception:
            pass
        
        return None
    
    def extract_pie_chart_colors(self, series) -> List[str]:
        """
        提取饼图各片段的颜色。
        
        参数：
            series: openpyxl饼图系列对象
            
        返回：
            颜色列表
        """
        colors = []
        
        try:
            # 检查数据点的个别颜色设置
            if hasattr(series, 'dPt') and series.dPt:
                for data_point in series.dPt:
                    color = self._extract_data_point_color(data_point)
                    if color:
                        colors.append(color)
            
            # 如果没有找到个别颜色，使用系列默认颜色
            if not colors:
                series_color = self.extract_series_color(series)
                if series_color:
                    point_count = self._get_series_point_count(series)
                    colors = generate_pie_color_variants(series_color, point_count)
        
        except Exception:
            pass
        
        return colors
    
    def extract_axis_title(self, title_obj: Any) -> Optional[str]:
        """
        安全提取openpyxl图表轴的标题。
        
        参数：
            title_obj: 图表轴的标题对象
            
        返回：
            标题字符串，提取失败时返回None
        """
        try:
            # 多种方式尝试提取标题文本
            if hasattr(title_obj, 'tx') and title_obj.tx:
                tx = title_obj.tx
                if hasattr(tx, 'rich') and tx.rich:
                    rich = tx.rich
                    if hasattr(rich, 'p') and rich.p and len(rich.p) > 0:
                        p = rich.p[0]
                        if hasattr(p, 'r') and p.r and hasattr(p.r, 't'):
                            return str(p.r.t)
            
            # 回退为字符串表示
            return str(title_obj) if title_obj else None
        except Exception:
            return None
    
    def _extract_from_val(self, val_obj) -> List[float]:
        """从val对象提取数值数据。"""
        data = []
        
        if hasattr(val_obj, 'numRef') and val_obj.numRef:
            try:
                if hasattr(val_obj.numRef, 'numCache') and val_obj.numRef.numCache:
                    # 从缓存中获取数据
                    data = [float(pt.v) for pt in val_obj.numRef.numCache.pt if pt.v is not None]
                else:
                    # 从引用中获取数据
                    data = [float(p.v) for p in val_obj.numRef.get_rows() if p.v is not None]
            except (AttributeError, TypeError, ValueError):
                pass
        
        return data
    
    def _extract_from_cat(self, cat_obj) -> List[str]:
        """从cat对象提取分类数据。"""
        data = []
        
        if hasattr(cat_obj, 'strRef') and cat_obj.strRef:
            try:
                if hasattr(cat_obj.strRef, 'strCache') and cat_obj.strRef.strCache:
                    # 从缓存中获取数据
                    data = [str(pt.v) for pt in cat_obj.strRef.strCache.pt if pt.v is not None]
                else:
                    # 从引用中获取数据
                    data = [str(p.v) for p in cat_obj.strRef.get_rows() if p.v is not None]
            except (AttributeError, TypeError):
                pass
        elif hasattr(cat_obj, 'numRef') and cat_obj.numRef:
            try:
                if hasattr(cat_obj.numRef, 'numCache') and cat_obj.numRef.numCache:
                    data = [str(pt.v) for pt in cat_obj.numRef.numCache.pt if pt.v is not None]
                else:
                    data = [str(p.v) for p in cat_obj.numRef.get_rows() if p.v is not None]
            except (AttributeError, TypeError):
                pass
        
        return data
    
    def _extract_from_xval(self, xval_obj) -> List[str]:
        """从xVal对象提取数据。"""
        data = []
        
        if hasattr(xval_obj, 'numRef') and xval_obj.numRef:
            try:
                data = [str(p.v) for p in xval_obj.numRef.get_rows() if p.v is not None]
            except (AttributeError, TypeError):
                pass
        elif hasattr(xval_obj, 'strRef') and xval_obj.strRef:
            try:
                data = [str(p.v) for p in xval_obj.strRef.get_rows() if p.v is not None]
            except (AttributeError, TypeError):
                pass
        
        return data
    
    def _extract_from_graphics_properties(self, series) -> Optional[str]:
        """从图形属性中提取颜色。"""
        if not (hasattr(series, 'graphicalProperties') and series.graphicalProperties):
            return None
        
        graphic_props = series.graphicalProperties
        
        # 检查solidFill
        if hasattr(graphic_props, 'solidFill') and graphic_props.solidFill:
            solid_fill = graphic_props.solidFill
            
            # RGB颜色
            if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr:
                rgb = solid_fill.srgbClr.val
                if rgb:
                    return f"#{rgb.upper()}"
            
            # 主题颜色
            if hasattr(solid_fill, 'schemeClr') and solid_fill.schemeClr:
                scheme_color = solid_fill.schemeClr.val
                return convert_scheme_color_to_hex(scheme_color)
        
        # 检查线条属性中的颜色
        if hasattr(graphic_props, 'ln') and graphic_props.ln:
            line = graphic_props.ln
            if hasattr(line, 'solidFill') and line.solidFill:
                solid_fill = line.solidFill
                if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr:
                    rgb = solid_fill.srgbClr.val
                    if rgb:
                        return f"#{rgb.upper()}"
        
        return None
    
    def _extract_from_sp_pr(self, series) -> Optional[str]:
        """从spPr属性中提取颜色。"""
        if not (hasattr(series, 'spPr') and series.spPr):
            return None
        
        sp_pr = series.spPr
        if hasattr(sp_pr, 'solidFill') and sp_pr.solidFill:
            solid_fill = sp_pr.solidFill
            if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr:
                rgb = solid_fill.srgbClr.val
                if rgb:
                    return f"#{rgb.upper()}"
            if hasattr(solid_fill, 'schemeClr') and solid_fill.schemeClr:
                scheme_color = solid_fill.schemeClr.val
                return convert_scheme_color_to_hex(scheme_color)
        
        return None
    
    def _extract_data_point_color(self, data_point) -> Optional[str]:
        """提取数据点的颜色。"""
        try:
            if hasattr(data_point, 'spPr') and data_point.spPr:
                sp_pr = data_point.spPr
                if hasattr(sp_pr, 'solidFill') and sp_pr.solidFill:
                    solid_fill = sp_pr.solidFill
                    if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr:
                        rgb = solid_fill.srgbClr.val
                        if rgb:
                            return f"#{rgb.upper()}"
                    if hasattr(solid_fill, 'schemeClr') and solid_fill.schemeClr:
                        scheme_color = solid_fill.schemeClr.val
                        return convert_scheme_color_to_hex(scheme_color)
        except Exception:
            pass
        
        return None
    
    def _get_series_point_count(self, series) -> int:
        """获取系列的数据点数量。"""
        try:
            if (hasattr(series, 'val') and series.val and 
                hasattr(series.val, 'numRef') and series.val.numRef and
                hasattr(series.val.numRef, 'numCache') and series.val.numRef.numCache):
                return len(series.val.numRef.numCache.pt)
        except Exception:
            pass
        
        return 3  # 默认数量