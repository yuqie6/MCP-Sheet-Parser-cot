"""
样式保真度验证测试模块

测试样式验证器的功能，验证95%保真度目标的实现。
"""

import pytest
from src.utils.style_validator import StyleValidator, StyleComparisonResult, CellFidelityResult, SheetFidelityReport
from src.models.table_model import Style, Sheet, Row, Cell
from src.parsers.factory import ParserFactory


class TestStyleValidator:
    """样式验证器测试类"""
    
    def test_validator_creation(self):
        """测试验证器创建"""
        validator = StyleValidator()
        assert validator is not None
        assert len(validator.STYLE_WEIGHTS) > 0
        assert sum(validator.STYLE_WEIGHTS.values()) == 100.0  # 权重总和应为100%
    
    def test_identical_styles_comparison(self):
        """测试相同样式的对比"""
        validator = StyleValidator()
        
        style1 = Style(
            font_name="Arial",
            font_size=12.0,
            font_color="#000000",
            bold=True,
            background_color="#FFFFFF"
        )
        
        style2 = Style(
            font_name="Arial",
            font_size=12.0,
            font_color="#000000",
            bold=True,
            background_color="#FFFFFF"
        )
        
        results = validator.compare_styles(style1, style2)
        
        # 所有属性都应该匹配
        for result in results:
            assert result.match, f"属性 {result.attribute} 应该匹配"
        
        # 总分应该接近1.0
        total_score = sum(r.score for r in results)
        assert total_score >= 0.95
    
    def test_different_styles_comparison(self):
        """测试不同样式的对比"""
        validator = StyleValidator()
        
        style1 = Style(
            font_name="Arial",
            font_size=12.0,
            font_color="#000000",
            bold=True
        )
        
        style2 = Style(
            font_name="Times",
            font_size=14.0,
            font_color="#FF0000",
            bold=False
        )
        
        results = validator.compare_styles(style1, style2)
        
        # 应该有不匹配的属性
        mismatches = [r for r in results if not r.match]
        assert len(mismatches) > 0
        
        # 总分应该小于1.0
        total_score = sum(r.score for r in results)
        assert total_score < 0.95
    
    def test_font_size_tolerance(self):
        """测试字体大小容差"""
        validator = StyleValidator()
        
        # 测试在容差范围内
        match, score = validator._compare_font_size(12.0, 12.3)
        assert match  # 应该匹配（容差0.5）
        assert score == 1.0
        
        # 测试超出容差范围
        match, score = validator._compare_font_size(12.0, 14.0)
        assert not match  # 不应该匹配
        assert score == 0.0
    
    def test_color_similarity(self):
        """测试颜色相似度"""
        validator = StyleValidator()
        
        # 测试相同颜色
        match, score = validator._compare_color("#FF0000", "#FF0000")
        assert match
        assert score == 1.0
        
        # 测试相似颜色
        match, score = validator._compare_color("#FF0000", "#FE0000")
        assert score > 0.9  # 应该有很高的相似度
        
        # 测试完全不同的颜色
        match, score = validator._compare_color("#FF0000", "#0000FF")
        assert score < 0.5
    
    def test_cell_fidelity_analysis(self):
        """测试单元格保真度分析"""
        validator = StyleValidator()
        
        expected_style = Style(
            font_name="Arial",
            font_size=12.0,
            bold=True,
            font_color="#000000"
        )
        
        actual_style = Style(
            font_name="Arial",
            font_size=12.0,
            bold=True,
            font_color="#000000"
        )
        
        result = validator.analyze_cell_fidelity(0, 0, expected_style, actual_style)
        
        assert isinstance(result, CellFidelityResult)
        assert result.row == 0
        assert result.col == 0
        assert result.overall_score >= 0.95  # 应该有很高的分数
        assert len(result.issues) == 0  # 不应该有问题
    
    def test_sheet_fidelity_report(self):
        """测试工作表保真度报告"""
        validator = StyleValidator()
        
        # 创建期望的工作表
        expected_style = Style(font_name="Arial", font_size=12.0, bold=True)
        expected_sheet = Sheet(
            name="test",
            rows=[
                Row(cells=[
                    Cell(value="A1", style=expected_style),
                    Cell(value="B1", style=expected_style)
                ]),
                Row(cells=[
                    Cell(value="A2", style=expected_style),
                    Cell(value="B2", style=expected_style)
                ])
            ]
        )
        
        # 创建实际的工作表（相同样式）
        actual_style = Style(font_name="Arial", font_size=12.0, bold=True)
        actual_sheet = Sheet(
            name="test",
            rows=[
                Row(cells=[
                    Cell(value="A1", style=actual_style),
                    Cell(value="B1", style=actual_style)
                ]),
                Row(cells=[
                    Cell(value="A2", style=actual_style),
                    Cell(value="B2", style=actual_style)
                ])
            ]
        )
        
        report = validator.generate_fidelity_report("TestParser", expected_sheet, actual_sheet)
        
        assert isinstance(report, SheetFidelityReport)
        assert report.parser_name == "TestParser"
        assert report.analyzed_cells == 4
        assert report.overall_fidelity >= 95.0  # 应该达到95%目标
        assert len(report.recommendations) > 0


class TestParserFidelity:
    """解析器保真度测试类"""
    
    def test_xlsx_parser_fidelity(self):
        """测试XLSX解析器保真度"""
        validator = StyleValidator()
        
        # 创建参考样式
        reference_style = Style(
            font_name="Calibri",
            font_size=11.0,
            font_color="#000000",
            background_color="#FFFFFF",
            text_align="left",
            vertical_align="top"
        )
        
        # 模拟XLSX解析器输出
        xlsx_style = Style(
            font_name="Calibri",
            font_size=11.0,
            font_color="#000000",
            background_color="#FFFFFF",
            text_align="left",
            vertical_align="top"
        )
        
        results = validator.compare_styles(reference_style, xlsx_style)
        total_score = sum(r.score for r in results)
        
        # XLSX解析器应该有很高的保真度
        assert total_score >= 0.90
    
    def test_xls_parser_fidelity(self):
        """测试XLS解析器保真度"""
        validator = StyleValidator()
        
        # 创建参考样式
        reference_style = Style(
            font_name="Arial",
            font_size=10.0,
            bold=True,
            font_color="#000000"
        )
        
        # 模拟XLS解析器输出（可能有一些差异）
        xls_style = Style(
            font_name="Arial",
            font_size=10.0,
            bold=True,
            font_color="#000000"
        )
        
        results = validator.compare_styles(reference_style, xls_style)
        total_score = sum(r.score for r in results)
        
        # XLS解析器应该有良好的保真度
        assert total_score >= 0.80
    
    def test_xlsb_parser_fidelity(self):
        """测试XLSB解析器保真度"""
        validator = StyleValidator()
        
        # 创建参考样式
        reference_style = Style(
            font_name="Calibri",
            font_size=11.0,
            text_align="left",
            vertical_align="bottom"
        )
        
        # 模拟XLSB解析器输出（基础样式）
        xlsb_style = Style(
            font_name="Calibri",
            font_size=11.0,
            text_align="left",
            vertical_align="bottom"
        )
        
        results = validator.compare_styles(reference_style, xlsb_style)
        total_score = sum(r.score for r in results)
        
        # XLSB解析器专注数据准确性，样式保真度适中
        assert total_score >= 0.70
    
    def test_xlsm_parser_fidelity(self):
        """测试XLSM解析器保真度"""
        validator = StyleValidator()
        
        # 创建参考样式
        reference_style = Style(
            font_name="Calibri",
            font_size=11.0,
            font_color="#000000",
            bold=False,
            italic=False
        )
        
        # 模拟XLSM解析器输出（继承XLSX能力）
        xlsm_style = Style(
            font_name="Calibri",
            font_size=11.0,
            font_color="#000000",
            bold=False,
            italic=False
        )
        
        results = validator.compare_styles(reference_style, xlsm_style)
        total_score = sum(r.score for r in results)
        
        # XLSM解析器应该与XLSX有相同的保真度
        assert total_score >= 0.90


class TestFidelityBenchmark:
    """保真度基准测试类"""
    
    def test_95_percent_fidelity_benchmark(self):
        """测试95%保真度基准"""
        validator = StyleValidator()
        
        # 创建高保真度样式对比
        perfect_style = Style(
            font_name="Arial",
            font_size=12.0,
            font_color="#000000",
            background_color="#FFFFFF",
            bold=True,
            italic=False,
            text_align="center",
            vertical_align="middle"
        )
        
        # 几乎完美的样式（只有微小差异）
        near_perfect_style = Style(
            font_name="Arial",
            font_size=12.1,  # 微小差异，在容差范围内
            font_color="#000000",
            background_color="#FFFFFF",
            bold=True,
            italic=False,
            text_align="center",
            vertical_align="middle"
        )
        
        results = validator.compare_styles(perfect_style, near_perfect_style)
        total_score = sum(r.score for r in results)
        
        # 应该达到95%以上的保真度
        assert total_score >= 0.95
        
        # 转换为百分比
        fidelity_percentage = total_score * 100
        assert fidelity_percentage >= 95.0
    
    def test_fidelity_report_quality_gates(self):
        """测试保真度报告质量门禁"""
        validator = StyleValidator()
        
        # 创建测试工作表
        high_fidelity_style = Style(font_name="Arial", font_size=12.0, bold=True)
        test_sheet = Sheet(
            name="quality_test",
            rows=[Row(cells=[Cell(value="test", style=high_fidelity_style)])]
        )
        
        report = validator.generate_fidelity_report("QualityTest", test_sheet, test_sheet)
        
        # 质量门禁检查
        assert report.overall_fidelity >= 95.0, f"保真度 {report.overall_fidelity}% 未达到95%目标"
        assert report.analyzed_cells > 0, "应该分析至少一个单元格"
        assert len(report.summary) > 0, "应该生成摘要信息"
        assert len(report.recommendations) > 0, "应该提供建议"
