"""
图表数据提取工具模块

提供统一的图表数据提取功能，避免代码重复。
"""

from typing import Any
from src.utils.color_utils import convert_scheme_color_to_hex, generate_pie_color_variants
from src.constants import ChartConstants


class ChartDataExtractor:
    """图表数据提取器，统一处理各种图表数据提取逻辑。"""
    
    def extract_series_y_data(self, series) -> list[float]:
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

    def extract_series_x_data(self, series) -> list[str]:
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

    def extract_series_color(self, series) -> str | None:
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

    def extract_pie_chart_colors(self, series) -> list[str]:
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

    def extract_axis_title(self, title_obj: Any) -> str | None:
        """
        安全提取openpyxl图表轴的标题。

        参数：
            title_obj: 图表轴的标题对象（可能是Title对象或Text对象）

        返回：
            标题字符串，提取失败时返回None
        """
        if not title_obj:
            return None

        # 尝试多种提取方法
        extraction_methods = [
            self._extract_from_title_tx,
            self._extract_from_rich_text,
            self._extract_from_string_reference,
            self._extract_from_direct_attributes,
            self._extract_from_string_representation
        ]

        for method in extraction_methods:
            try:
                result = method(title_obj)
                if result:
                    return result
            except Exception:
                continue

        return None

    def _extract_from_title_tx(self, title_obj: Any) -> str | None:
        """从Title对象的tx属性提取文本"""
        if hasattr(title_obj, 'tx') and title_obj.tx:
            return self._extract_text_from_tx(title_obj.tx)
        return None

    def _extract_from_rich_text(self, title_obj: Any) -> str | None:
        """从rich文本对象提取文本"""
        if hasattr(title_obj, 'rich') and title_obj.rich:
            return self._extract_text_from_rich(title_obj.rich)
        return None

    def _extract_from_string_reference(self, title_obj: Any) -> str | None:
        """从字符串引用对象提取文本"""
        if hasattr(title_obj, 'strRef') and title_obj.strRef:
            return self._extract_text_from_strref(title_obj.strRef)
        return None

    def _extract_from_direct_attributes(self, title_obj: Any) -> str | None:
        """从直接属性提取文本"""
        # 检查常见的文本属性
        for attr in ['text', 'value', 'v']:
            if hasattr(title_obj, attr):
                attr_value = getattr(title_obj, attr)
                if attr_value:
                    text_str = str(attr_value).strip()
                    if text_str and not text_str.startswith('<'):
                        return text_str
        return None

    def _extract_from_string_representation(self, title_obj: Any) -> str | None:
        """从字符串表示提取文本"""
        try:
            str_repr = str(title_obj)
            if (str_repr and
                not str_repr.startswith('<') and
                not 'object at' in str_repr and
                not 'Parameters:' in str_repr and
                len(str_repr) < ChartConstants.MAX_STRING_REPRESENTATION_LENGTH):
                return str_repr.strip()
        except (AttributeError, TypeError, ValueError):
            pass
        return None
    
    def _extract_text_from_tx(self, tx) -> str | None:
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

    def _extract_text_from_rich(self, rich) -> str | None:
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

    def _extract_text_from_strref(self, strref) -> str | None:
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

    def extract_color(self, solid_fill) -> str | None:
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
    
    def _extract_from_val(self, val_obj) -> list[float]:
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
    
    def _extract_from_cat(self, cat_obj) -> list[str]:
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

    def _extract_from_xval(self, xval_obj) -> list[str]:
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

    def _extract_from_graphics_properties(self, series) -> str | None:
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

    def _extract_from_sp_pr(self, series) -> str | None:
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

    def _extract_data_point_color(self, data_point) -> str | None:
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

        return ChartConstants.DEFAULT_PIE_SLICE_COUNT  # 默认数量

    def extract_data_labels(self, series) -> dict[str, Any]:
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
            'show_category': False,  
            'show_category_name': False,
            'show_series': False,
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
                    show_cat = bool(dLbls.showCatName)
                    data_labels_info['show_category_name'] = show_cat
                    data_labels_info['show_category'] = show_cat

                if hasattr(dLbls, 'showSerName') and dLbls.showSerName is not None:
                    show_ser = bool(dLbls.showSerName)
                    data_labels_info['show_series_name'] = show_ser
                    data_labels_info['show_series'] = show_ser

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

    def _extract_individual_data_label(self, dLbl) -> dict[str, Any] | None:
        """提取单个数据标签的信息。"""
        try:
            label_info: dict[str, Any] = {
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

    def extract_chart_title(self, chart) -> str | None:
        """
        提取图表的主标题。
        此方法利用现有的 extract_axis_title，因为标题结构相似。
        """
        if hasattr(chart, 'title') and chart.title:
            return self.extract_axis_title(chart.title)
        return None

    def extract_plot_area(self, chart) -> dict[str, Any]:
        """
        提取图表绘图区域的背景色和布局信息。
        """
        plot_area_info: dict[str, Any] = {
            'background_color': None,
            'layout': None
        }
        if not (hasattr(chart, 'plotArea') and chart.plotArea):
            return plot_area_info

        plot_area = chart.plotArea

        # 提取背景颜色
        if hasattr(plot_area, 'spPr') and plot_area.spPr and hasattr(plot_area.spPr, 'solidFill'):
            plot_area_info['background_color'] = self.extract_color(plot_area.spPr.solidFill)

        # 提取布局信息
        if hasattr(plot_area, 'layout') and plot_area.layout and hasattr(plot_area.layout, 'manualLayout'):
            manual_layout = plot_area.layout.manualLayout
            plot_area_info['layout'] = {
                'x': getattr(manual_layout, 'x', None),
                'y': getattr(manual_layout, 'y', None),
                'width': getattr(manual_layout, 'w', None),
                'height': getattr(manual_layout, 'h', None)
            }
            
        return plot_area_info

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

            # 提取文本框和形状注释
            textbox_annotations = self._extract_textbox_annotations(chart)
            annotations.extend(textbox_annotations)

            # 提取形状注释
            shape_annotations = self._extract_shape_annotations(chart)
            annotations.extend(shape_annotations)

            # 提取绘图区域中的其他注释元素
            plotarea_annotations = self._extract_plotarea_annotations(chart)
            annotations.extend(plotarea_annotations)

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract chart annotations: {e}")

        return annotations

    def _extract_textbox_annotations(self, chart) -> list[dict[str, Any]]:
        """
        提取图表中的文本框注释。

        注意：由于openpyxl对文本框的支持有限，此方法尝试多种方式来提取文本框信息。

        参数：
            chart: openpyxl图表对象

        返回：
            文本框注释列表
        """
        textbox_annotations = []

        try:
            # 方法1：尝试从plotArea访问文本框
            if hasattr(chart, 'plotArea') and chart.plotArea:
                plotarea = chart.plotArea

                # 检查plotArea中的文本属性
                if hasattr(plotarea, 'txPr') and plotarea.txPr:
                    text_content = self._extract_text_from_rich(plotarea.txPr)
                    if text_content:
                        textbox_annotations.append({
                            'type': 'textbox',
                            'text': text_content,
                            'position': 'plotarea',
                            'source': 'plotArea.txPr'
                        })

                # 检查plotArea中可能的文本元素
                for attr_name in ['tx', 'text', 'textBox', 'txBox']:
                    if hasattr(plotarea, attr_name):
                        attr_value = getattr(plotarea, attr_name)
                        if attr_value:
                            text_content = self.extract_axis_title(attr_value)
                            if text_content:
                                textbox_annotations.append({
                                    'type': 'textbox',
                                    'text': text_content,
                                    'position': 'plotarea',
                                    'source': f'plotArea.{attr_name}'
                                })

            # 方法2：尝试从图表的其他属性访问文本框
            # 检查图表级别的文本属性
            for attr_name in ['textBox', 'txBox', 'freeText', 'annotation']:
                if hasattr(chart, attr_name):
                    attr_value = getattr(chart, attr_name)
                    if attr_value:
                        # 如果是列表，遍历每个元素
                        if isinstance(attr_value, (list, tuple)):
                            for i, item in enumerate(attr_value):
                                text_content = self.extract_axis_title(item)
                                if text_content:
                                    textbox_annotations.append({
                                        'type': 'textbox',
                                        'text': text_content,
                                        'position': 'chart',
                                        'source': f'chart.{attr_name}[{i}]'
                                    })
                        else:
                            text_content = self.extract_axis_title(attr_value)
                            if text_content:
                                textbox_annotations.append({
                                    'type': 'textbox',
                                    'text': text_content,
                                    'position': 'chart',
                                    'source': f'chart.{attr_name}'
                                })

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract textbox annotations: {e}")

        return textbox_annotations

    def _extract_shape_annotations(self, chart) -> list[dict[str, Any]]:
        """
        提取图表中的形状注释。

        注意：openpyxl对形状的支持非常有限。此方法尝试从可能的属性中提取形状信息。

        参数：
            chart: openpyxl图表对象

        返回：
            形状注释列表
        """
        shape_annotations = []

        try:
            # 方法1：尝试从plotArea访问形状
            if hasattr(chart, 'plotArea') and chart.plotArea:
                plotarea = chart.plotArea

                # 检查plotArea中的形状属性
                for attr_name in ['sp', 'shape', 'shapes', 'spPr']:
                    if hasattr(plotarea, attr_name):
                        attr_value = getattr(plotarea, attr_name)
                        if attr_value:
                            shapes_found = self._extract_shapes_from_attribute(attr_value, f'plotArea.{attr_name}')
                            shape_annotations.extend(shapes_found)

            # 方法2：尝试从图表级别访问形状
            for attr_name in ['shapes', 'sp', 'drawing', 'drawingElements']:
                if hasattr(chart, attr_name):
                    attr_value = getattr(chart, attr_name)
                    if attr_value:
                        shapes_found = self._extract_shapes_from_attribute(attr_value, f'chart.{attr_name}')
                        shape_annotations.extend(shapes_found)

            # 方法3：尝试从图表的图形属性中提取形状信息
            if hasattr(chart, 'graphicalProperties') and chart.graphicalProperties:
                gp = chart.graphicalProperties
                shape_info = self._extract_shape_from_graphical_properties(gp)
                if shape_info:
                    shape_annotations.append(shape_info)

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract shape annotations: {e}")

        return shape_annotations

    def _extract_shapes_from_attribute(self, attr_value, source: str) -> list[dict[str, Any]]:
        """从属性值中提取形状信息。"""
        shapes = []

        try:
            # 如果是列表或元组，遍历每个元素
            if isinstance(attr_value, (list, tuple)):
                for i, item in enumerate(attr_value):
                    shape_info = self._extract_single_shape(item, f'{source}[{i}]')
                    if shape_info:
                        shapes.append(shape_info)
            else:
                # 单个形状对象
                shape_info = self._extract_single_shape(attr_value, source)
                if shape_info:
                    shapes.append(shape_info)

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract shapes from {source}: {e}")

        return shapes

    def _extract_single_shape(self, shape_obj, source: str) -> dict[str, Any] | None:
        """
        提取单个形状的信息。
        重构此函数以提高健壮性和清晰度。
        只有在找到文本内容时才认为它是一个有效的注释。
        """
        try:
            # 首先，尝试提取文本内容
            text_content = None
            for text_attr in ['text', 'tx', 'txPr', 'textBody']:
                if hasattr(shape_obj, text_attr):
                    text_value = getattr(shape_obj, text_attr)
                    if text_value:
                        # 使用现有的轴标题提取器，因为它能处理复杂的文本结构
                        extracted_text = self.extract_axis_title(text_value)
                        if extracted_text and not extracted_text.startswith('<'):
                            text_content = extracted_text
                            break
            
            # 如果没有文本内容，则认为这不是一个注释形状，直接返回None
            if not text_content:
                return None

            # 如果有文本内容，则构建形状信息字典
            shape_info = {
                'type': 'shape',
                'position': 'unknown',
                'source': source,
                'text': text_content,
                'shape_type': None,
                'properties': {}
            }

            # 提取形状类型
            if hasattr(shape_obj, 'type'):
                shape_info['shape_type'] = str(shape_obj.type)
            elif hasattr(shape_obj, '__class__'):
                shape_info['shape_type'] = shape_obj.__class__.__name__

            # 提取其他形状属性
            if hasattr(shape_obj, 'spPr') and shape_obj.spPr:
                shape_info['properties']['spPr'] = 'present'

            return shape_info

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract single shape from {source}: {e}")

        return None

    def _extract_shape_from_graphical_properties(self, gp) -> dict[str, Any] | None:
        """从图形属性中提取形状信息。"""
        try:
            shape_info = {
                'type': 'shape',
                'position': 'chart',
                'source': 'chart.graphicalProperties',
                'text': None,
                'shape_type': 'graphical_properties',
                'properties': {}
            }

            # 检查是否有填充属性
            if hasattr(gp, 'solidFill') and gp.solidFill:
                shape_info['properties']['fill'] = 'solid'
            elif hasattr(gp, 'noFill') and gp.noFill:
                shape_info['properties']['fill'] = 'none'

            # 检查是否有线条属性
            if hasattr(gp, 'ln') and gp.ln:
                shape_info['properties']['line'] = 'present'

            # 只有在找到属性时才返回
            if shape_info['properties']:
                return shape_info

        except Exception:
            pass

        return None

    def _extract_plotarea_annotations(self, chart) -> list[dict[str, Any]]:
        """
        提取绘图区域中的其他注释元素。

        这个方法尝试从图表的绘图区域中提取各种可能的注释元素，
        包括但不限于：数据标签、趋势线标签、误差线等。

        参数：
            chart: openpyxl图表对象

        返回：
            绘图区域注释列表
        """
        plotarea_annotations = []

        try:
            if not (hasattr(chart, 'plotArea') and chart.plotArea):
                return plotarea_annotations

            plotarea = chart.plotArea

            # 方法1：检查绘图区域的布局信息
            if hasattr(plotarea, 'layout') and plotarea.layout:
                layout_info = self._extract_layout_annotations(plotarea.layout)
                if layout_info:
                    plotarea_annotations.append(layout_info)

            # 方法2：检查绘图区域中的数据表
            if hasattr(plotarea, 'dTable') and plotarea.dTable:
                table_info = {
                    'type': 'data_table',
                    'position': 'plotarea',
                    'source': 'plotArea.dTable',
                    'enabled': True
                }
                plotarea_annotations.append(table_info)

            # 方法3：检查绘图区域中的其他可能的注释元素
            annotation_attrs = [
                'annotation', 'annotations', 'textAnnotation',
                'callout', 'callouts', 'note', 'notes',
                'label', 'labels', 'marker', 'markers'
            ]

            for attr_name in annotation_attrs:
                if hasattr(plotarea, attr_name):
                    attr_value = getattr(plotarea, attr_name)
                    if attr_value:
                        annotations_found = self._extract_annotations_from_attribute(
                            attr_value, f'plotArea.{attr_name}'
                        )
                        plotarea_annotations.extend(annotations_found)

            # 方法4：尝试从XML属性中提取未知的注释元素
            # 这是一个实验性的方法，尝试发现openpyxl可能没有完全支持的元素
            if hasattr(plotarea, '__dict__'):
                for attr_name, attr_value in plotarea.__dict__.items():
                    if (attr_name.startswith('_') or
                        attr_name in ['tagname', 'namespace'] or
                        attr_value is None):
                        continue

                    # 检查是否是可能包含文本的属性
                    if any(keyword in attr_name.lower() for keyword in
                           ['text', 'annotation', 'label', 'note', 'comment']):
                        text_content = self._try_extract_text_from_unknown_element(attr_value)
                        if text_content:
                            plotarea_annotations.append({
                                'type': 'unknown_annotation',
                                'text': text_content,
                                'position': 'plotarea',
                                'source': f'plotArea.{attr_name}',
                                'experimental': True
                            })

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract plotarea annotations: {e}")

        return plotarea_annotations

    def _extract_layout_annotations(self, layout) -> dict[str, Any] | None:
        """从布局对象中提取注释信息。"""
        try:
            if not layout:
                return None

            layout_info = {
                'type': 'layout',
                'position': 'plotarea',
                'source': 'plotArea.layout',
                'properties': {}
            }

            # 提取布局属性
            layout_attrs = ['x', 'y', 'w', 'h', 'xMode', 'yMode', 'wMode', 'hMode']
            for attr in layout_attrs:
                if hasattr(layout, attr):
                    value = getattr(layout, attr)
                    if value is not None:
                        layout_info['properties'][attr] = str(value)

            # 只有在找到属性时才返回
            if layout_info['properties']:
                return layout_info

        except Exception:
            pass

        return None

    def _extract_annotations_from_attribute(self, attr_value, source: str) -> list[dict[str, Any]]:
        """从属性值中提取注释信息。"""
        annotations = []

        try:
            # 如果是列表或元组，遍历每个元素
            if isinstance(attr_value, (list, tuple)):
                for i, item in enumerate(attr_value):
                    annotation_info = self._extract_single_annotation(item, f'{source}[{i}]')
                    if annotation_info:
                        annotations.append(annotation_info)
            else:
                # 单个注释对象
                annotation_info = self._extract_single_annotation(attr_value, source)
                if annotation_info:
                    annotations.append(annotation_info)

        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.debug(f"Failed to extract annotations from {source}: {e}")

        return annotations

    def _extract_single_annotation(self, annotation_obj, source: str) -> dict[str, Any] | None:
        """提取单个注释的信息。"""
        # 如果输入为None，直接返回None
        if annotation_obj is None:
            return None

        try:
            annotation_info = {
                'type': 'annotation',
                'position': 'plotarea',
                'source': source,
                'text': None,
                'properties': {}
            }

            # 尝试提取文本内容
            text_content = self.extract_axis_title(annotation_obj)
            if text_content:
                annotation_info['text'] = text_content

            # 尝试提取其他属性
            if hasattr(annotation_obj, '__class__'):
                annotation_info['properties']['class'] = annotation_obj.__class__.__name__

            # 只有在找到文本或属性时才返回
            if annotation_info['text'] or annotation_info['properties']:
                return annotation_info

        except Exception:
            pass

        return None

    def _try_extract_text_from_unknown_element(self, element) -> str | None:
        """尝试从未知元素中提取文本内容。"""
        try:
            # 尝试多种方法提取文本
            if isinstance(element, str):
                return element.strip() if element.strip() else None

            # 尝试使用现有的文本提取方法
            text_content = self.extract_axis_title(element)
            if text_content:
                return text_content

            # 如果是对象，尝试查找文本属性
            if hasattr(element, '__dict__'):
                for attr_name, attr_value in element.__dict__.items():
                    if 'text' in attr_name.lower() and isinstance(attr_value, str):
                        text = attr_value.strip()
                        if text:
                            return text

        except Exception:
            pass

        return None