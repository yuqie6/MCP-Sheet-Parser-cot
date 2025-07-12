"""
样式保真度验证工具模块

提供样式对比、保真度评分和质量验证功能，确保解析器达到95%样式保真度目标。
"""

import logging
from typing import Dict, List, Tuple, Any
from dataclasses import dataclass
from src.models.table_model import Style, Sheet, Row, Cell

logger = logging.getLogger(__name__)


@dataclass
class StyleComparisonResult:
    """样式对比结果"""
    attribute: str
    expected: Any
    actual: Any
    match: bool
    weight: float
    score: float


@dataclass
class CellFidelityResult:
    """单元格保真度结果"""
    row: int
    col: int
    style_comparisons: List[StyleComparisonResult]
    overall_score: float
    issues: List[str]


@dataclass
class SheetFidelityReport:
    """工作表保真度报告"""
    parser_name: str
    total_cells: int
    analyzed_cells: int
    overall_fidelity: float
    cell_results: List[CellFidelityResult]
    summary: Dict[str, Any]
    recommendations: List[str]


class StyleValidator:
    """样式保真度验证器"""
    
    # 样式属性权重配置（总和为100%）
    STYLE_WEIGHTS = {
        # 字体属性 (40%)
        'font_name': 8.0,
        'font_size': 8.0,
        'font_color': 8.0,
        'bold': 5.0,
        'italic': 5.0,
        'underline': 6.0,
        
        # 背景和填充 (25%)
        'background_color': 25.0,
        
        # 对齐方式 (15%)
        'text_align': 8.0,
        'vertical_align': 7.0,
        
        # 边框 (15%)
        'border_top': 3.75,
        'border_bottom': 3.75,
        'border_left': 3.75,
        'border_right': 3.75,
        
        # 其他属性 (5%)
        'wrap_text': 2.0,
        'number_format': 3.0
    }
    
    def __init__(self):
        """初始化样式验证器"""
        self.tolerance_config = {
            'font_size_tolerance': 0.5,  # 字体大小容差
            'color_similarity_threshold': 0.9,  # 颜色相似度阈值
            'ignore_default_values': True  # 忽略默认值比较
        }
    
    def compare_styles(self, expected: Style, actual: Style) -> List[StyleComparisonResult]:
        """
        对比两个样式对象，返回详细的对比结果。
        
        Args:
            expected: 期望的样式
            actual: 实际的样式
            
        Returns:
            样式对比结果列表
        """
        results = []
        
        for attr, weight in self.STYLE_WEIGHTS.items():
            expected_value = getattr(expected, attr, None)
            actual_value = getattr(actual, attr, None)
            
            # 执行属性特定的比较
            match, score = self._compare_attribute(attr, expected_value, actual_value)
            
            results.append(StyleComparisonResult(
                attribute=attr,
                expected=expected_value,
                actual=actual_value,
                match=match,
                weight=weight,
                score=score * weight / 100.0
            ))
        
        return results
    
    def _compare_attribute(self, attr: str, expected: Any, actual: Any) -> Tuple[bool, float]:
        """
        比较单个样式属性。
        
        Args:
            attr: 属性名
            expected: 期望值
            actual: 实际值
            
        Returns:
            (是否匹配, 得分0-1)
        """
        # 处理None值
        if expected is None and actual is None:
            return True, 1.0
        if expected is None or actual is None:
            # 如果忽略默认值，且其中一个是默认值，则认为匹配
            if self.tolerance_config['ignore_default_values']:
                if self._is_default_value(attr, expected) or self._is_default_value(attr, actual):
                    return True, 1.0
            return False, 0.0
        
        # 字体大小特殊处理
        if attr == 'font_size':
            return self._compare_font_size(expected, actual)
        
        # 颜色特殊处理
        if 'color' in attr:
            return self._compare_color(expected, actual)
        
        # 边框特殊处理
        if 'border_' in attr:
            return self._compare_border(expected, actual)
        
        # 默认精确匹配
        match = expected == actual
        return match, 1.0 if match else 0.0
    
    def _compare_font_size(self, expected: float, actual: float) -> Tuple[bool, float]:
        """比较字体大小，允许容差"""
        if not isinstance(expected, (int, float)) or not isinstance(actual, (int, float)):
            return expected == actual, 1.0 if expected == actual else 0.0
        
        diff = abs(float(expected) - float(actual))
        tolerance = self.tolerance_config['font_size_tolerance']
        
        if diff <= tolerance:
            return True, 1.0
        elif diff <= tolerance * 2:
            return False, 0.5  # 部分匹配
        else:
            return False, 0.0
    
    def _compare_color(self, expected: str, actual: str) -> Tuple[bool, float]:
        """比较颜色，支持相似度匹配"""
        if expected == actual:
            return True, 1.0
        
        # 简化的颜色相似度计算
        if isinstance(expected, str) and isinstance(actual, str):
            if expected.startswith('#') and actual.startswith('#'):
                similarity = self._calculate_color_similarity(expected, actual)
                threshold = self.tolerance_config['color_similarity_threshold']
                
                if similarity >= threshold:
                    return True, similarity
                else:
                    return False, similarity
        
        return False, 0.0
    
    def _compare_border(self, expected: str, actual: str) -> Tuple[bool, float]:
        """比较边框样式"""
        # 简化边框比较：如果都为空或都非空，给予部分分数
        expected_empty = not expected or expected == ""
        actual_empty = not actual or actual == ""
        
        if expected_empty and actual_empty:
            return True, 1.0
        elif expected_empty != actual_empty:
            return False, 0.0
        else:
            # 都非空，检查是否完全匹配
            return expected == actual, 1.0 if expected == actual else 0.7
    
    def _calculate_color_similarity(self, color1: str, color2: str) -> float:
        """计算颜色相似度（简化版本）"""
        try:
            # 移除#号
            c1 = color1.lstrip('#')
            c2 = color2.lstrip('#')
            
            # 转换为RGB
            r1, g1, b1 = int(c1[0:2], 16), int(c1[2:4], 16), int(c1[4:6], 16)
            r2, g2, b2 = int(c2[0:2], 16), int(c2[2:4], 16), int(c2[4:6], 16)
            
            # 计算欧几里得距离
            distance = ((r1-r2)**2 + (g1-g2)**2 + (b1-b2)**2) ** 0.5
            max_distance = (255**2 * 3) ** 0.5
            
            similarity = 1.0 - (distance / max_distance)
            return max(0.0, similarity)
        except:
            return 0.0
    
    def _is_default_value(self, attr: str, value: Any) -> bool:
        """检查是否为默认值"""
        default_values = {
            'font_name': None,
            'font_size': None,
            'font_color': "#000000",
            'background_color': "#FFFFFF",
            'text_align': "left",
            'vertical_align': "top",
            'bold': False,
            'italic': False,
            'underline': False,
            'wrap_text': False,
            'number_format': "",
            'border_top': "",
            'border_bottom': "",
            'border_left': "",
            'border_right': ""
        }
        
        return value == default_values.get(attr)
    
    def analyze_cell_fidelity(self, row: int, col: int, expected: Style, actual: Style) -> CellFidelityResult:
        """
        分析单个单元格的样式保真度。
        
        Args:
            row: 行索引
            col: 列索引
            expected: 期望样式
            actual: 实际样式
            
        Returns:
            单元格保真度结果
        """
        comparisons = self.compare_styles(expected, actual)
        overall_score = sum(comp.score for comp in comparisons)
        
        # 生成问题列表
        issues = []
        for comp in comparisons:
            if not comp.match and comp.weight > 5.0:  # 只报告重要属性的问题
                issues.append(f"{comp.attribute}: 期望 {comp.expected}, 实际 {comp.actual}")
        
        return CellFidelityResult(
            row=row,
            col=col,
            style_comparisons=comparisons,
            overall_score=overall_score,
            issues=issues
        )
    
    def generate_fidelity_report(self, parser_name: str, expected_sheet: Sheet, actual_sheet: Sheet) -> SheetFidelityReport:
        """
        生成工作表的完整保真度报告。
        
        Args:
            parser_name: 解析器名称
            expected_sheet: 期望的工作表
            actual_sheet: 实际解析的工作表
            
        Returns:
            保真度报告
        """
        cell_results = []
        total_score = 0.0
        analyzed_cells = 0
        
        # 分析每个单元格
        max_rows = min(len(expected_sheet.rows), len(actual_sheet.rows))
        for row_idx in range(max_rows):
            expected_row = expected_sheet.rows[row_idx]
            actual_row = actual_sheet.rows[row_idx]
            
            max_cols = min(len(expected_row.cells), len(actual_row.cells))
            for col_idx in range(max_cols):
                expected_cell = expected_row.cells[col_idx]
                actual_cell = actual_row.cells[col_idx]
                
                if expected_cell.style and actual_cell.style:
                    result = self.analyze_cell_fidelity(
                        row_idx, col_idx, expected_cell.style, actual_cell.style
                    )
                    cell_results.append(result)
                    total_score += result.overall_score
                    analyzed_cells += 1
        
        # 计算总体保真度
        overall_fidelity = (total_score / analyzed_cells * 100) if analyzed_cells > 0 else 0.0
        
        # 生成摘要
        summary = self._generate_summary(cell_results, overall_fidelity)
        
        # 生成建议
        recommendations = self._generate_recommendations(cell_results, overall_fidelity)
        
        return SheetFidelityReport(
            parser_name=parser_name,
            total_cells=max_rows * max_cols if max_rows > 0 and max_cols > 0 else 0,
            analyzed_cells=analyzed_cells,
            overall_fidelity=overall_fidelity,
            cell_results=cell_results,
            summary=summary,
            recommendations=recommendations
        )
    
    def _generate_summary(self, cell_results: List[CellFidelityResult], overall_fidelity: float) -> Dict[str, Any]:
        """生成保真度摘要"""
        if not cell_results:
            return {}
        
        # 统计各属性的平均得分
        attr_scores = {}
        for result in cell_results:
            for comp in result.style_comparisons:
                if comp.attribute not in attr_scores:
                    attr_scores[comp.attribute] = []
                attr_scores[comp.attribute].append(comp.score * 100)
        
        attr_averages = {attr: sum(scores)/len(scores) for attr, scores in attr_scores.items()}
        
        # 找出问题最多的属性
        problem_attrs = sorted(attr_averages.items(), key=lambda x: x[1])[:3]
        
        return {
            'overall_fidelity': overall_fidelity,
            'attribute_scores': attr_averages,
            'top_problem_attributes': problem_attrs,
            'cells_with_issues': len([r for r in cell_results if r.issues]),
            'perfect_cells': len([r for r in cell_results if r.overall_score >= 0.95])
        }
    
    def _generate_recommendations(self, cell_results: List[CellFidelityResult], overall_fidelity: float) -> List[str]:
        """生成优化建议"""
        recommendations = []
        
        if overall_fidelity < 95.0:
            recommendations.append(f"当前保真度 {overall_fidelity:.1f}% 未达到95%目标，需要优化")
        
        if overall_fidelity >= 95.0:
            recommendations.append(f"恭喜！保真度 {overall_fidelity:.1f}% 已达到95%目标")
        elif overall_fidelity >= 90.0:
            recommendations.append("保真度接近目标，重点优化主要样式属性")
        elif overall_fidelity >= 80.0:
            recommendations.append("保真度良好，需要系统性优化")
        else:
            recommendations.append("保真度较低，需要全面检查解析器实现")
        
        return recommendations
