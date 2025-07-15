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
            title_obj: 图表轴的标题对象（可能是Title对象或Text对象）
            
        返回：
            标题字符串，提取失败时返回None
        """
        if not title_obj:
            return None
        
        # 记录对象类型以供调试
        obj_type = type(title_obj).__name__
        
        try:
            # 方式1：处理Title对象 - 从 title.tx.rich.p[0].r[0].t 提取
            if hasattr(title_obj, 'tx') and title_obj.tx:
                tx = title_obj.tx
                result = self._extract_text_from_tx(tx)
                if result:
                    return result
            
            # 方式2：处理Text对象 - 直接从 rich.p[0].r[0].t 提取
            if hasattr(title_obj, 'rich') and title_obj.rich:
                result = self._extract_text_from_rich(title_obj.rich)
                if result:
                    return result
            
            # 方式3：处理含有strRef的对象
            if hasattr(title_obj, 'strRef') and title_obj.strRef:
                result = self._extract_text_from_strref(title_obj.strRef)
                if result:
                    return result
            
            # 方式4：检查是否有直接的文本属性
            if hasattr(title_obj, 'text') and title_obj.text:
                text_str = str(title_obj.text).strip()
                if text_str and not text_str.startswith('<'):  # 避免对象字符串
                    return text_str
                
            # 方式5：检查是否有value属性
            if hasattr(title_obj, 'value') and title_obj.value:
                value_str = str(title_obj.value).strip()
                if value_str and not value_str.startswith('<'):  # 避免对象字符串
                    return value_str
            
            # 方式6：检查是否有v属性（某些对象直接存储文本）
            if hasattr(title_obj, 'v') and title_obj.v:
                v_str = str(title_obj.v).strip()
                if v_str and not v_str.startswith('<'):  # 避免对象字符串
                    return v_str
            
            # 方式7：如果对象有__str__方法但不是默认的对象表示，尝试使用
            try:
                str_repr = str(title_obj)
                if (str_repr and 
                    not str_repr.startswith('<') and 
                    not 'object at' in str_repr and 
                    not 'Parameters:' in str_repr and
                    len(str_repr) < 100):  # 避免长对象描述
                    return str_repr.strip()
            except:
                pass
            
            # 如果以上都失败，返回None
            return None
            
        except Exception as e:
            # 在调试模式下记录错误
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract title from {obj_type}: {e}")
            return None
    
    def _extract_text_from_tx(self, tx) -> Optional[str]:
        """从tx对象提取文本内容"""
        if not tx:
            return None
            
        # 从rich文本提取
        if hasattr(tx, 'rich') and tx.rich:
            text = self._extract_text_from_rich(tx.rich)
            if text:
                return text
        
        # 从strRef提取
        if hasattr(tx, 'strRef') and tx.strRef:
            return self._extract_text_from_strref(tx.strRef)
        
        # 直接文本属性
        if hasattr(tx, 'v') and tx.v:
            v_str = str(tx.v).strip()
            if v_str:
                return v_str
        
        return None
    
    def _extract_text_from_rich(self, rich) -> Optional[str]:
        """从rich对象提取文本内容"""
        if not rich or not hasattr(rich, 'p'):
            return None
        
        # 遍历所有段落
        for p in rich.p:
            if hasattr(p, 'r') and p.r:
                # 收集所有run中的文本
                texts = []
                for run in p.r:
                    if hasattr(run, 't') and run.t:
                        text_content = str(run.t).strip()
                        if text_content:
                            texts.append(text_content)
                
                if texts:
                    return ''.join(texts)
        
        return None
    
    def _extract_text_from_strref(self, strref) -> Optional[str]:
        """从strRef对象提取文本内容"""
        if not strref:
            return None
        
        # 从缓存中提取
        if hasattr(strref, 'strCache') and strref.strCache:
            if hasattr(strref.strCache, 'pt') and strref.strCache.pt:
                for pt in strref.strCache.pt:
                    if hasattr(pt, 'v') and pt.v:
                        v_str = str(pt.v).strip()
                        if v_str:
                            return v_str
        
        # 从公式中提取
        if hasattr(strref, 'f') and strref.f:
            f_str = str(strref.f).strip()
            if f_str:
                return f_str
        
        return None

    def extract_color(self, solid_fill) -> Optional[str]:
        """
        从solidFill对象中提取颜色。
        
        参数：
            solid_fill: openpyxl的solidFill对象
            
        返回：
            颜色的十六进制字符串，提取失败时返回None
        """
        if hasattr(solid_fill, 'srgbClr') and solid_fill.srgbClr and solid_fill.srgbClr.val:
            return f"#{solid_fill.srgbClr.val.upper()}"
        if hasattr(solid_fill, 'schemeClr') and solid_fill.schemeClr and solid_fill.schemeClr.val:
            return convert_scheme_color_to_hex(solid_fill.schemeClr.val)
        return None
    
    def _extract_from_val(self, val_obj) -> List[float]:
        """从val对象提取数值数据。"""
        data = []
        
        if hasattr(val_obj, 'numRef') and val_obj.numRef:
            try:
                if hasattr(val_obj.numRef, 'numCache') and val_obj.numRef.numCache:
                    # 从缓存中获取数据
                    data = [float(pt.v) for pt in val_obj.numRef.numCache.pt if pt.v is not None]
                elif hasattr(val_obj.numRef, 'get_rows'):
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
                elif hasattr(cat_obj.strRef, 'get_rows'):
                    # 从引用中获取数据
                    data = [str(p.v) for p in cat_obj.strRef.get_rows() if p.v is not None]
            except (AttributeError, TypeError):
                pass
        elif hasattr(cat_obj, 'numRef') and cat_obj.numRef:
            try:
                if hasattr(cat_obj.numRef, 'numCache') and cat_obj.numRef.numCache:
                    data = [str(pt.v) for pt in cat_obj.numRef.numCache.pt if pt.v is not None]
                elif hasattr(cat_obj.numRef, 'get_rows'):
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

    def extract_data_labels(self, series) -> Dict[str, Any]:
        """
        提取系列的数据标签信息。

        参数：
            series: openpyxl图表系列对象

        返回：
            包含数据标签信息的字典
        """
        data_labels_info = {
            'enabled': False,
            'show_value': False,
            'show_category_name': False,
            'show_series_name': False,
            'show_percent': False,
            'show_legend_key': False,
            'show_leader_lines': False,
            'position': None,
            'separator': None,
            'number_format': None,
            'labels': []  # 个别数据标签
        }

        try:
            # 检查系列级别的数据标签设置
            if hasattr(series, 'dLbls') and series.dLbls:
                dLbls = series.dLbls
                data_labels_info['enabled'] = True

                # 提取显示选项
                if hasattr(dLbls, 'showVal') and dLbls.showVal is not None:
                    data_labels_info['show_value'] = bool(dLbls.showVal)

                if hasattr(dLbls, 'showCatName') and dLbls.showCatName is not None:
                    data_labels_info['show_category_name'] = bool(dLbls.showCatName)

                if hasattr(dLbls, 'showSerName') and dLbls.showSerName is not None:
                    data_labels_info['show_series_name'] = bool(dLbls.showSerName)

                if hasattr(dLbls, 'showPercent') and dLbls.showPercent is not None:
                    data_labels_info['show_percent'] = bool(dLbls.showPercent)

                if hasattr(dLbls, 'showLegendKey') and dLbls.showLegendKey is not None:
                    data_labels_info['show_legend_key'] = bool(dLbls.showLegendKey)

                if hasattr(dLbls, 'showLeaderLines') and dLbls.showLeaderLines is not None:
                    data_labels_info['show_leader_lines'] = bool(dLbls.showLeaderLines)

                # 提取位置信息
                if hasattr(dLbls, 'dLblPos') and dLbls.dLblPos:
                    data_labels_info['position'] = str(dLbls.dLblPos)

                # 提取分隔符
                if hasattr(dLbls, 'separator') and dLbls.separator:
                    data_labels_info['separator'] = str(dLbls.separator)

                # 提取数字格式
                if hasattr(dLbls, 'numFmt') and dLbls.numFmt:
                    data_labels_info['number_format'] = str(dLbls.numFmt)

                # 提取个别数据标签
                if hasattr(dLbls, 'dLbl') and dLbls.dLbl:
                    for dLbl in dLbls.dLbl:
                        label_info = self._extract_individual_data_label(dLbl)
                        if label_info:
                            data_labels_info['labels'].append(label_info)

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract data labels: {e}")

        return data_labels_info

    def _extract_individual_data_label(self, dLbl) -> Optional[Dict[str, Any]]:
        """提取单个数据标签的信息。"""
        try:
            label_info = {
                'index': None,
                'text': None,
                'position': None,
                'show_value': None,
                'show_category_name': None,
                'show_series_name': None,
                'show_percent': None
            }

            # 提取索引
            if hasattr(dLbl, 'idx') and dLbl.idx is not None:
                label_info['index'] = int(dLbl.idx)

            # 提取文本内容
            if hasattr(dLbl, 'txPr') and dLbl.txPr:
                text = self._extract_text_from_rich(dLbl.txPr)
                if text:
                    label_info['text'] = text

            # 提取位置
            if hasattr(dLbl, 'dLblPos') and dLbl.dLblPos:
                label_info['position'] = str(dLbl.dLblPos)

            # 提取显示选项
            if hasattr(dLbl, 'showVal') and dLbl.showVal is not None:
                label_info['show_value'] = bool(dLbl.showVal)

            if hasattr(dLbl, 'showCatName') and dLbl.showCatName is not None:
                label_info['show_category_name'] = bool(dLbl.showCatName)

            if hasattr(dLbl, 'showSerName') and dLbl.showSerName is not None:
                label_info['show_series_name'] = bool(dLbl.showSerName)

            if hasattr(dLbl, 'showPercent') and dLbl.showPercent is not None:
                label_info['show_percent'] = bool(dLbl.showPercent)

            return label_info

        except Exception:
            return None

    def extract_legend_info(self, chart) -> dict[str, Any]:
        """
        提取图表的图例信息。

        参数：
            chart: openpyxl图表对象

        返回：
            包含图例信息的字典
        """
        legend_info = {
            'enabled': False,
            'position': None,
            'overlay': False,
            'entries': []
        }

        try:
            if hasattr(chart, 'legend') and chart.legend:
                legend = chart.legend
                legend_info['enabled'] = True

                # 提取位置信息
                if hasattr(legend, 'position') and legend.position:
                    legend_info['position'] = str(legend.position)

                # 提取是否覆盖图表
                if hasattr(legend, 'overlay') and legend.overlay is not None:
                    legend_info['overlay'] = bool(legend.overlay)

                # 提取图例条目
                if hasattr(legend, 'legendEntry') and legend.legendEntry:
                    for entry in legend.legendEntry:
                        entry_info = self._extract_legend_entry(entry)
                        if entry_info:
                            legend_info['entries'].append(entry_info)

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract legend info: {e}")

        return legend_info

    def _extract_legend_entry(self, entry) -> dict[str, Any] | None:
        """提取单个图例条目的信息。"""
        try:
            entry_info = {
                'index': None,
                'text': None,
                'delete': False
            }

            # 提取索引
            if hasattr(entry, 'idx') and entry.idx is not None:
                entry_info['index'] = int(entry.idx)

            # 提取文本 - 尝试多种方式
            text = None

            # 方式1：从txPr提取
            if hasattr(entry, 'txPr') and entry.txPr:
                text = self._extract_text_from_rich(entry.txPr)

            # 方式2：从tx提取
            if not text and hasattr(entry, 'tx') and entry.tx:
                text = self.extract_axis_title(entry.tx)

            # 方式3：从其他可能的文本属性提取
            if not text:
                for attr in ['text', 'value', 'v']:
                    if hasattr(entry, attr):
                        attr_value = getattr(entry, attr)
                        if attr_value:
                            text = str(attr_value).strip()
                            if text and not text.startswith('<'):
                                break
                            text = None

            if text:
                entry_info['text'] = text

            # 检查是否被删除
            if hasattr(entry, 'delete') and entry.delete is not None:
                entry_info['delete'] = bool(entry.delete)

            return entry_info

        except Exception:
            return None

    def extract_chart_annotations(self, chart) -> list[dict[str, Any]]:
        """
        提取图表的注释和文本框信息。

        参数：
            chart: openpyxl图表对象

        返回：
            包含注释信息的列表
        """
        annotations = []

        try:
            # 检查图表标题作为注释
            if hasattr(chart, 'title') and chart.title:
                title_text = self.extract_axis_title(chart.title)
                if title_text:
                    annotations.append({
                        'type': 'title',
                        'text': title_text,
                        'position': 'top'
                    })

            # 检查轴标题作为注释
            if hasattr(chart, 'x_axis') and chart.x_axis and chart.x_axis.title:
                x_title = self.extract_axis_title(chart.x_axis.title)
                if x_title:
                    annotations.append({
                        'type': 'axis_title',
                        'text': x_title,
                        'position': 'bottom',
                        'axis': 'x'
                    })

            if hasattr(chart, 'y_axis') and chart.y_axis and chart.y_axis.title:
                y_title = self.extract_axis_title(chart.y_axis.title)
                if y_title:
                    annotations.append({
                        'type': 'axis_title',
                        'text': y_title,
                        'position': 'left',
                        'axis': 'y'
                    })

            # TODO: 添加对其他类型注释的支持（如文本框、形状等）
            # 这需要访问图表的绘图区域和形状集合

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract chart annotations: {e}")

        return annotations