"""
性能优化功能测试模块

测试性能分析、文件大小估算和智能处理建议功能。
"""

import pytest
from src.utils.performance import PerformanceOptimizer, PerformanceMetrics
from src.models.table_model import Sheet, Row, Cell, Style


class TestPerformanceOptimizer:
    """性能优化器测试类。"""
    
    def test_performance_optimizer_initialization(self):
        """测试性能优化器初始化。"""
        optimizer = PerformanceOptimizer()
        
        assert optimizer.start_time is None
        assert optimizer.SMALL_FILE_THRESHOLD == 1000
        assert optimizer.MEDIUM_FILE_THRESHOLD == 10000
        assert optimizer.LARGE_FILE_THRESHOLD == 50000
    
    def test_timing_functionality(self):
        """测试计时功能。"""
        optimizer = PerformanceOptimizer()
        
        # 测试开始计时
        optimizer.start_timing()
        assert optimizer.start_time is not None
        
        # 测试结束计时
        elapsed = optimizer.end_timing()
        assert isinstance(elapsed, float)
        assert elapsed >= 0
    
    def test_small_sheet_analysis(self):
        """测试小表格的性能分析。"""
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()
        
        # 创建小表格（3x3）
        cells = [
            [Cell(value="A1"), Cell(value="B1"), Cell(value="C1")],
            [Cell(value="A2"), Cell(value="B2"), Cell(value="C2")],
            [Cell(value="A3"), Cell(value="B3"), Cell(value="C3")]
        ]
        rows = [Row(cells=row) for row in cells]
        sheet = Sheet(name="test", rows=rows)
        
        # 分析性能
        metrics = optimizer.analyze_sheet_performance(sheet)
        
        # 验证结果
        assert metrics.rows == 3
        assert metrics.cols == 3
        assert metrics.total_cells == 9
        assert metrics.non_empty_cells == 9
        assert metrics.recommendation == "small_file_direct"
        assert isinstance(metrics.estimated_html_size, int)
        assert isinstance(metrics.estimated_json_size, int)
    
    def test_styled_cells_analysis(self):
        """测试有样式单元格的分析。"""
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()
        
        # 创建有样式的单元格
        styled_cell = Cell(
            value="Styled",
            style=Style(bold=True, font_color="#FF0000")
        )
        normal_cell = Cell(value="Normal")
        
        rows = [Row(cells=[styled_cell, normal_cell])]
        sheet = Sheet(name="test", rows=rows)
        
        # 分析性能
        metrics = optimizer.analyze_sheet_performance(sheet)
        
        # 验证样式单元格统计
        assert metrics.styled_cells == 1
        assert metrics.non_empty_cells == 2
    
    def test_hyperlink_and_comment_analysis(self):
        """测试超链接和注释单元格的分析。"""
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()
        
        # 创建有超链接和注释的单元格
        hyperlink_cell = Cell(
            value="Link",
            style=Style(hyperlink="https://example.com")
        )
        comment_cell = Cell(
            value="Comment",
            style=Style(comment="This is a comment")
        )
        both_cell = Cell(
            value="Both",
            style=Style(
                hyperlink="https://test.com",
                comment="Link with comment"
            )
        )
        
        rows = [Row(cells=[hyperlink_cell, comment_cell, both_cell])]
        sheet = Sheet(name="test", rows=rows)
        
        # 分析性能
        metrics = optimizer.analyze_sheet_performance(sheet)
        
        # 验证进阶功能统计
        assert metrics.styled_cells == 3  # 所有单元格都有样式
    
    def test_large_file_recommendation(self):
        """测试大文件的处理建议。"""
        optimizer = PerformanceOptimizer()
        
        # 创建大表格（模拟）
        large_sheet = Sheet(name="large", rows=[])
        # 模拟大文件的指标
        large_metrics = PerformanceMetrics(
            rows=2000,
            cols=50,
            total_cells=100000,  # 超过大文件阈值
            non_empty_cells=80000,
            styled_cells=10000,
            estimated_html_size=500000,
            estimated_json_size=300000,
            processing_time=2.5,
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
    
    def test_pagination_calculation(self):
        """测试分页参数计算。"""
        optimizer = PerformanceOptimizer()
        
        # 创建测试表格
        rows = [Row(cells=[Cell(value=f"Row{i}")]) for i in range(2500)]
        sheet = Sheet(name="test", rows=rows)
        
        # 计算分页参数
        pagination = optimizer.calculate_pagination_params(sheet, max_rows_per_page=1000)
        
        # 验证分页结果
        assert pagination["total_rows"] == 2500
        assert pagination["max_rows_per_page"] == 1000
        assert pagination["total_pages"] == 3
        assert len(pagination["pages"]) == 3
        
        # 验证第一页
        first_page = pagination["pages"][0]
        assert first_page["page_number"] == 1
        assert first_page["start_row"] == 0
        assert first_page["end_row"] == 1000
        assert first_page["row_count"] == 1000
        
        # 验证最后一页
        last_page = pagination["pages"][-1]
        assert last_page["page_number"] == 3
        assert last_page["start_row"] == 2000
        assert last_page["end_row"] == 2500
        assert last_page["row_count"] == 500
    
    def test_has_custom_style_detection(self):
        """测试自定义样式检测。"""
        optimizer = PerformanceOptimizer()
        
        # 测试默认样式
        default_style = Style()
        assert optimizer._has_custom_style(default_style) == False
        
        # 测试自定义样式
        custom_style = Style(bold=True)
        assert optimizer._has_custom_style(custom_style) == True
        
        # 测试超链接样式
        hyperlink_style = Style(hyperlink="https://example.com")
        assert optimizer._has_custom_style(hyperlink_style) == True
        
        # 测试注释样式
        comment_style = Style(comment="Test comment")
        assert optimizer._has_custom_style(comment_style) == True
        
        # 测试None样式
        assert optimizer._has_custom_style(None) == False
    
    def test_size_estimation_accuracy(self):
        """测试大小估算的合理性。"""
        optimizer = PerformanceOptimizer()
        optimizer.start_timing()
        
        # 创建测试表格
        cells = [[Cell(value=f"Cell{i}{j}") for j in range(5)] for i in range(10)]
        rows = [Row(cells=row) for row in cells]
        sheet = Sheet(name="test", rows=rows)
        
        # 分析性能
        metrics = optimizer.analyze_sheet_performance(sheet)
        
        # 验证估算合理性
        assert metrics.estimated_html_size > 0
        assert metrics.estimated_json_size > 0
        assert metrics.estimated_html_size > metrics.estimated_json_size  # HTML通常比JSON大
        
        # 验证估算与实际数据的关系
        expected_content_size = sum(len(str(cell.value)) for row in sheet.rows for cell in row.cells)
        assert metrics.estimated_html_size > expected_content_size  # 包含标签开销
        assert metrics.estimated_json_size > expected_content_size  # 包含JSON结构开销


class TestPerformanceMetrics:
    """性能指标数据类测试。"""
    
    def test_performance_metrics_creation(self):
        """测试性能指标对象创建。"""
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
