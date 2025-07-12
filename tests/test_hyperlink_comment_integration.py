"""
超链接和注释功能的集成测试

测试从Excel文件解析到HTML转换的完整流程
"""

import pytest
import tempfile
import os
import openpyxl
from pathlib import Path
from src.core_service import CoreService


class TestHyperlinkCommentIntegration:
    """超链接和注释功能的集成测试。"""
    
    def test_xlsx_hyperlink_comment_end_to_end(self):
        """测试XLSX文件中超链接和注释的端到端处理。"""
        # 创建包含超链接和注释的XLSX文件
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        
        # 添加普通数据
        worksheet['A1'] = 'Name'
        worksheet['B1'] = 'Website'
        worksheet['C1'] = 'Notes'
        
        # 添加带超链接的数据
        worksheet['A2'] = 'Google'
        worksheet['B2'] = 'Visit Google'
        worksheet['B2'].hyperlink = 'https://www.google.com'
        
        # 添加带注释的数据
        worksheet['C2'] = 'Search Engine'
        comment = openpyxl.comments.Comment('This is the most popular search engine', 'Test Author')
        worksheet['C2'].comment = comment
        
        # 添加同时包含超链接和注释的数据
        worksheet['A3'] = 'GitHub'
        worksheet['B3'] = 'Visit GitHub'
        worksheet['B3'].hyperlink = 'https://github.com'
        github_comment = openpyxl.comments.Comment('Code hosting platform', 'Test Author')
        worksheet['B3'].comment = github_comment
        
        # 保存到临时文件
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as xlsx_file:
            xlsx_path = xlsx_file.name
        
        workbook.save(xlsx_path)
        workbook.close()
        
        try:
            service = CoreService()
            
            # 1. 测试解析功能
            print(f"\n=== 测试解析功能 ===")
            json_data = service.parse_sheet(xlsx_path)
            
            # 验证解析结果包含超链接和注释信息
            print(f"解析到的行数: {len(json_data['rows'])}")
            print(f"表头: {json_data.get('headers', [])}")

            # JSON格式中headers和rows是分开的，rows不包含表头
            assert len(json_data['rows']) == 2  # 数据行，不包括表头

            # 检查第一行数据（Google）
            google_row = json_data['rows'][0]  # 第一行数据
            website_cell = google_row[1]  # B2单元格
            
            # 验证超链接信息在JSON中
            print(f"Google网站单元格: {website_cell}")
            # 注意：JSON格式可能不直接包含样式信息，这取决于具体实现
            
            # 2. 测试HTML转换功能
            print(f"\n=== 测试HTML转换功能 ===")
            with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as html_file:
                html_path = html_file.name
            
            try:
                html_result = service.convert_to_html(xlsx_path, html_path)
                
                # 验证转换成功
                assert html_result['status'] == 'success'
                assert os.path.exists(html_path)
                
                # 读取HTML内容
                with open(html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                print(f"HTML内容长度: {len(html_content)} 字符")
                
                # 验证超链接转换
                assert 'href="https://www.google.com"' in html_content
                assert 'href="https://github.com"' in html_content
                assert 'Visit Google' in html_content
                assert 'Visit GitHub' in html_content
                
                # 验证注释转换
                assert 'title="This is the most popular search engine"' in html_content
                assert 'title="Code hosting platform"' in html_content
                
                # 验证HTML结构
                assert '<table>' in html_content
                assert '<a href=' in html_content  # 包含超链接
                assert 'title=' in html_content    # 包含工具提示
                
                print(f"✅ 超链接和注释功能集成测试成功")
                
            finally:
                if os.path.exists(html_path):
                    os.unlink(html_path)
            
        finally:
            os.unlink(xlsx_path)
    
    def test_csv_no_hyperlink_comment_support(self):
        """测试CSV文件不支持超链接和注释（确保不会出错）。"""
        # 创建简单的CSV文件
        csv_content = """Name,Website,Notes
Google,https://www.google.com,Search Engine
GitHub,https://github.com,Code Platform"""
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
            csv_file.write(csv_content)
            csv_path = csv_file.name
        
        try:
            service = CoreService()
            
            # 测试CSV解析（不应该有超链接和注释）
            json_data = service.parse_sheet(csv_path)
            assert len(json_data['rows']) == 2
            
            # 测试HTML转换（应该正常工作）
            with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as html_file:
                html_path = html_file.name
            
            try:
                html_result = service.convert_to_html(csv_path, html_path)
                assert html_result['status'] == 'success'
                
                # 读取HTML内容
                with open(html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                # CSV中的URL应该作为普通文本显示，不是超链接
                assert 'https://www.google.com' in html_content
                assert 'href=' not in html_content  # 不应该有超链接
                assert 'title=' not in html_content or 'title="表格:' in html_content  # 不应该有注释工具提示
                
                print(f"✅ CSV文件处理正常，无超链接和注释")
                
            finally:
                if os.path.exists(html_path):
                    os.unlink(html_path)
            
        finally:
            os.unlink(csv_path)
