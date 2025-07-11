"""
工具模块测试

测试工具模块的功能，包括性能优化器等。
"""

import pytest
from src.utils.performance import PerformanceOptimizer, PerformanceMetrics
from src.models.table_model import Sheet, Row, Cell, Style


class TestPerformanceOptimizerAdvanced:
    """性能优化器高级测试。"""
    
    def test_complex_sheet_analysis(self):
        """测试复杂表格的性能分析。"""
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()
        
        # 创建复杂表格
        rows = []
        for i in range(100):
            cells = []
            for j in range(20):
                # 创建不同类型的单元格
                if i == 0:  # 标题行
                    cell = Cell(
                        value=f"Column {j}",
                        style=Style(bold=True, background_color="#CCCCCC")
                    )
                elif j == 0:  # 第一列
                    cell = Cell(
                        value=f"Row {i}",
                        style=Style(italic=True)
                    )
                elif i % 10 == 0:  # 每10行添加超链接
                    cell = Cell(
                        value=f"Link {i}-{j}",
                        style=Style(hyperlink=f"https://example.com/{i}/{j}")
                    )
                elif j % 5 == 0:  # 每5列添加注释
                    cell = Cell(
                        value=f"Data {i}-{j}",
                        style=Style(comment=f"Comment for {i}-{j}")
                    )
                else:
                    cell = Cell(value=f"Data {i}-{j}")
                
                cells.append(cell)
            rows.append(Row(cells=cells))
        
        sheet = Sheet(name="complex_sheet", rows=rows)
        
        # 分析性能
        metrics = optimizer.analyze_sheet_performance(sheet)
        
        # 验证结果
        assert metrics.rows == 100
        assert metrics.cols == 20
        assert metrics.total_cells == 2000
        assert metrics.non_empty_cells == 2000
        assert metrics.styled_cells > 0  # 应该有样式单元格
        assert metrics.estimated_html_size > 0
        assert metrics.estimated_json_size > 0
        assert metrics.processing_time >= 0
    
    def test_empty_sheet_analysis(self):
        """测试空表格的性能分析。"""
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()
        
        # 创建空表格
        sheet = Sheet(name="empty", rows=[])
        
        # 分析性能
        metrics = optimizer.analyze_sheet_performance(sheet)
        
        # 验证结果
        assert metrics.rows == 0
        assert metrics.cols == 0
        assert metrics.total_cells == 0
        assert metrics.non_empty_cells == 0
        assert metrics.styled_cells == 0
        assert metrics.recommendation == "small_file_direct"
    
    def test_single_cell_sheet_analysis(self):
        """测试单单元格表格的性能分析。"""
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()
        
        # 创建单单元格表格
        cell = Cell(
            value="Single Cell",
            style=Style(
                bold=True,
                hyperlink="https://example.com",
                comment="This is the only cell"
            )
        )
        sheet = Sheet(name="single", rows=[Row(cells=[cell])])
        
        # 分析性能
        metrics = optimizer.analyze_sheet_performance(sheet)
        
        # 验证结果
        assert metrics.rows == 1
        assert metrics.cols == 1
        assert metrics.total_cells == 1
        assert metrics.non_empty_cells == 1
        assert metrics.styled_cells == 1
        assert metrics.recommendation == "small_file_direct"
    
    def test_large_sheet_pagination_recommendation(self):
        """测试大表格的分页建议。"""
        optimizer = PerformanceOptimizer()
        
        # 模拟大表格的指标
        large_metrics = PerformanceMetrics(
            rows=10000,
            cols=50,
            total_cells=500000,  # 超过大文件阈值
            non_empty_cells=400000,
            styled_cells=50000,
            estimated_html_size=2000000,
            estimated_json_size=1000000,
            processing_time=5.0,
            recommendation="very_large_file_require_pagination"
        )
        
        # 测试分页判断
        assert optimizer.should_use_pagination(large_metrics) == True
        
        # 测试建议信息
        message = optimizer.get_recommendation_message(
            large_metrics.recommendation, large_metrics
        )
        assert "超大文件" in message
        assert "分页处理" in message
        assert "500,000" in message  # 格式化的数字
    
    def test_medium_file_recommendations(self):
        """测试中等文件的建议。"""
        optimizer = PerformanceOptimizer()
        
        # 测试不同的中等文件场景
        scenarios = [
            {
                "total_cells": 5000,
                "html_size": 30000,
                "expected": "medium_file_direct"
            },
            {
                "total_cells": 8000,
                "html_size": 80000,
                "expected": "medium_file_with_warning"
            }
        ]
        
        for scenario in scenarios:
            recommendation = optimizer._get_processing_recommendation(
                scenario["total_cells"],
                scenario["html_size"],
                scenario["total_cells"] * 0.8  # 假设80%非空
            )
            assert recommendation == scenario["expected"]
    
    def test_size_estimation_scaling(self):
        """测试大小估算的扩展性。"""
        optimizer = PerformanceOptimizer()
        
        # 创建不同大小的表格并比较估算
        sizes = [10, 100, 1000]
        html_sizes = []
        json_sizes = []
        
        for size in sizes:
            # 创建指定大小的表格
            rows = []
            for i in range(size):
                cells = [Cell(value=f"Cell{i}{j}") for j in range(10)]
                rows.append(Row(cells=cells))
            
            sheet = Sheet(name=f"test_{size}", rows=rows)
            optimizer.start_timing()
            metrics = optimizer.analyze_sheet_performance(sheet)
            
            html_sizes.append(metrics.estimated_html_size)
            json_sizes.append(metrics.estimated_json_size)
        
        # 验证大小估算随表格大小增长
        assert html_sizes[0] < html_sizes[1] < html_sizes[2]
        assert json_sizes[0] < json_sizes[1] < json_sizes[2]
    
    def test_performance_metrics_dataclass(self):
        """测试性能指标数据类的功能。"""
        # 创建性能指标实例
        metrics = PerformanceMetrics(
            rows=100,
            cols=10,
            total_cells=1000,
            non_empty_cells=800,
            styled_cells=200,
            estimated_html_size=50000,
            estimated_json_size=30000,
            processing_time=1.5,
            recommendation="medium_file_direct"
        )
        
        # 验证所有属性
        assert metrics.rows == 100
        assert metrics.cols == 10
        assert metrics.total_cells == 1000
        assert metrics.non_empty_cells == 800
        assert metrics.styled_cells == 200
        assert metrics.estimated_html_size == 50000
        assert metrics.estimated_json_size == 30000
        assert metrics.processing_time == 1.5
        assert metrics.recommendation == "medium_file_direct"
        
        # 验证数据类的字符串表示
        metrics_str = str(metrics)
        assert "PerformanceMetrics" in metrics_str
        assert "rows=100" in metrics_str
    
    def test_optimizer_timing_edge_cases(self):
        """测试优化器计时的边界情况。"""
        optimizer = PerformanceOptimizer()
        
        # 测试未开始计时就结束
        elapsed = optimizer.end_timing()
        assert elapsed == 0.0
        
        # 测试多次开始计时
        optimizer.start_timing()
        first_start = optimizer.start_time
        optimizer.start_timing()
        second_start = optimizer.start_time
        
        assert second_start >= first_start
        
        # 测试正常计时
        elapsed = optimizer.end_timing()
        assert elapsed >= 0.0
    
    def test_custom_style_detection_edge_cases(self):
        """测试自定义样式检测的边界情况。"""
        optimizer = PerformanceOptimizer()
        
        # 测试各种样式组合
        test_cases = [
            (Style(), False),  # 默认样式
            (Style(bold=True), True),  # 单一样式
            (Style(font_color="#000000"), False),  # 默认颜色
            (Style(font_color="#FF0000"), True),  # 自定义颜色
            (Style(background_color="#FFFFFF"), False),  # 默认背景
            (Style(background_color="#FFFF00"), True),  # 自定义背景
            (Style(text_align="left"), False),  # 默认对齐
            (Style(text_align="center"), True),  # 自定义对齐
            (Style(hyperlink=""), False),  # 空超链接
            (Style(hyperlink="https://example.com"), True),  # 有效超链接
            (Style(comment=""), False),  # 空注释
            (Style(comment="Test comment"), True),  # 有效注释
            (None, False),  # None样式
        ]
        
        for style, expected in test_cases:
            result = optimizer._has_custom_style(style)
            assert result == expected, f"Failed for style: {style}"
