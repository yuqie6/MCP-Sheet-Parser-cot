"""
端到端完整测试模块

测试从文件输入到最终输出的完整工作流程
验证系统在真实使用场景下的表现
"""

import pytest
import tempfile
import os
import json
from pathlib import Path
from src.core_service import CoreService
from src.mcp_server.tools import (
    _handle_convert_to_html,
    _handle_parse_sheet,
    _handle_apply_changes
)


class TestCompleteWorkflow:
    """完整工作流程的端到端测试。"""
    
    def test_csv_complete_workflow(self):
        """测试CSV文件的完整处理工作流程。"""
        # 1. 创建测试数据
        csv_content = """Product,Price,Category,Stock
Laptop,999.99,Electronics,50
Mouse,29.99,Electronics,100
Book,19.99,Education,200
Chair,149.99,Furniture,25
Desk,299.99,Furniture,15"""
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
            csv_file.write(csv_content)
            csv_path = csv_file.name
        
        with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as html_file:
            html_path = html_file.name
        
        try:
            service = CoreService()
            
            # 2. 步骤1：获取文件信息
            print(f"\n=== 步骤1：获取文件信息 ===")
            info = service.get_sheet_info(csv_path)
            print(f"文件格式: {info['file_format']}")
            print(f"文件大小: {info['file_size']} bytes")
            print(f"数据维度: {info['dimensions']['rows']}行 x {info['dimensions']['cols']}列")
            print(f"总单元格数: {info['dimensions']['total_cells']}")
            
            assert info['file_format'] == '.csv'
            assert info['dimensions']['rows'] == 6  # 包括表头
            assert info['dimensions']['cols'] == 4
            assert info['dimensions']['total_cells'] == 24
            
            # 3. 步骤2：解析为JSON
            print(f"\n=== 步骤2：解析为JSON ===")
            json_data = service.parse_sheet(csv_path)
            print(f"表格名称: {json_data['sheet_name']}")
            print(f"表头: {json_data['headers']}")
            print(f"数据行数: {len(json_data['rows'])}")
            print(f"处理模式: {json_data['size_info']['processing_mode']}")
            
            assert json_data['sheet_name'] == Path(csv_path).stem
            assert json_data['headers'] == ['Product', 'Price', 'Category', 'Stock']
            assert len(json_data['rows']) == 5  # 不包括表头
            assert json_data['metadata']['total_rows'] == 6
            assert json_data['metadata']['total_cols'] == 4
            
            # 验证第一行数据
            first_row = json_data['rows'][0]
            assert first_row[0]['value'] == 'Laptop'
            assert first_row[1]['value'] == '999.99'
            assert first_row[2]['value'] == 'Electronics'
            assert first_row[3]['value'] == '50'
            
            # 4. 步骤3：转换为HTML
            print(f"\n=== 步骤3：转换为HTML ===")
            html_result = service.convert_to_html(csv_path, html_path)
            print(f"转换状态: {html_result['status']}")
            print(f"输出文件: {html_result['output_path']}")
            print(f"文件大小: {html_result['file_size']} bytes")
            print(f"转换行数: {html_result['rows_converted']}")
            print(f"转换单元格数: {html_result['cells_converted']}")
            
            assert html_result['status'] == 'success'
            assert html_result['rows_converted'] == 6
            assert html_result['cells_converted'] == 24
            assert os.path.exists(html_result['output_path'])
            
            # 5. 步骤4：验证HTML内容
            print(f"\n=== 步骤4：验证HTML内容 ===")
            with open(html_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
            
            print(f"HTML文件大小: {len(html_content)} 字符")
            
            # 验证HTML结构
            assert "<table>" in html_content
            assert "</table>" in html_content
            assert "<tr>" in html_content  # 实际使用简单的tr结构
            assert "<td>" in html_content
            
            # 验证数据内容
            assert "Product" in html_content
            assert "Laptop" in html_content
            assert "999.99" in html_content
            assert "Electronics" in html_content
            
            # 验证所有产品都在HTML中
            products = ['Laptop', 'Mouse', 'Book', 'Chair', 'Desk']
            for product in products:
                assert product in html_content
            
            # 6. 步骤5：数据修改测试
            print(f"\n=== 步骤5：数据修改测试 ===")
            
            # 修改数据（模拟价格调整）
            modified_json = json_data.copy()
            # 将Laptop价格从999.99改为899.99
            modified_json['rows'][0][1]['value'] = '899.99'
            
            modify_result = service.apply_changes(csv_path, modified_json, create_backup=True)
            print(f"修改状态: {modify_result['status']}")
            print(f"备份创建: {modify_result['backup_created']}")
            
            assert "验证成功" in modify_result['status'] or modify_result['status'] == 'success'
            assert modify_result['backup_created'] is True
            
            print(f"\n=== 工作流程完成 ===")
            print("✅ 所有步骤都成功执行")
            
        finally:
            # 清理文件
            os.unlink(csv_path)
            if os.path.exists(html_path):
                os.unlink(html_path)
            # 清理可能的备份文件
            backup_path = csv_path + ".backup"
            if os.path.exists(backup_path):
                os.unlink(backup_path)
    
    @pytest.mark.asyncio
    async def test_mcp_tools_complete_workflow(self):
        """测试通过MCP工具的完整工作流程。"""
        # 创建测试数据
        csv_content = "Name,Score,Grade\nAlice,85,B\nBob,92,A\nCharlie,78,C"
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
            csv_file.write(csv_content)
            csv_path = csv_file.name
        
        with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as html_file:
            html_path = html_file.name
        
        try:
            core_service = CoreService()
            
            print(f"\n=== MCP工具工作流程测试 ===")
            
            # 1. 使用parse_sheet工具
            print(f"\n--- 1. 解析文件 ---")
            parse_args = {"file_path": csv_path}
            parse_result = await _handle_parse_sheet(parse_args, core_service)
            
            assert isinstance(parse_result, list)
            parse_text = parse_result[0].text
            print(f"解析结果: {parse_text[:200]}...")
            
            # 验证JSON格式
            assert '"sheet_name"' in parse_text
            assert '"headers"' in parse_text
            assert '"rows"' in parse_text
            
            # 2. 使用convert_to_html工具
            print(f"\n--- 2. 转换为HTML ---")
            convert_args = {
                "file_path": csv_path,
                "output_path": html_path
            }
            convert_result = await _handle_convert_to_html(convert_args, core_service)
            
            assert isinstance(convert_result, list)
            convert_text = convert_result[0].text
            print(f"转换结果: {convert_text[:200]}...")
            
            # 验证转换成功
            assert '"status": "success"' in convert_text
            assert '"output_path"' in convert_text
            assert os.path.exists(html_path)
            
            # 3. 使用apply_changes工具
            print(f"\n--- 3. 应用数据修改 ---")
            
            # 先获取原始数据格式
            original_data = core_service.parse_sheet(csv_path)
            table_model_json = {
                "sheet_name": original_data["sheet_name"],
                "headers": original_data["headers"],
                "rows": original_data["rows"]
            }
            
            apply_args = {
                "file_path": csv_path,
                "table_model_json": table_model_json,
                "create_backup": True
            }
            apply_result = await _handle_apply_changes(apply_args, core_service)
            
            assert isinstance(apply_result, list)
            apply_text = apply_result[0].text
            print(f"修改结果: {apply_text[:200]}...")
            
            # 验证修改操作
            assert "验证成功" in apply_text or '"status": "success"' in apply_text
            
            # 4. 验证最终HTML文件
            print(f"\n--- 4. 验证最终结果 ---")
            with open(html_path, 'r', encoding='utf-8') as f:
                final_html = f.read()
            
            print(f"最终HTML大小: {len(final_html)} 字符")
            
            # 验证HTML包含所有数据
            assert "Alice" in final_html
            assert "Bob" in final_html
            assert "Charlie" in final_html
            assert "85" in final_html
            assert "92" in final_html
            assert "78" in final_html
            
            print(f"✅ MCP工具工作流程测试完成")
            
        finally:
            # 清理文件
            os.unlink(csv_path)
            if os.path.exists(html_path):
                os.unlink(html_path)
            backup_path = csv_path + ".backup"
            if os.path.exists(backup_path):
                os.unlink(backup_path)
    
    def test_large_file_handling_workflow(self):
        """测试大文件处理的完整工作流程。"""
        print(f"\n=== 大文件处理测试 ===")
        
        # 创建一个较大的CSV文件（1000行）
        header = "ID,Name,Email,Department,Salary,StartDate"
        rows = []
        for i in range(1000):
            row = f"{i+1},Employee{i+1},emp{i+1}@company.com,Dept{i%10},{50000+i*100},2023-{(i%12)+1:02d}-{(i%28)+1:02d}"
            rows.append(row)
        
        large_csv_content = header + "\n" + "\n".join(rows)
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
            csv_file.write(large_csv_content)
            csv_path = csv_file.name
        
        try:
            service = CoreService()
            
            print(f"创建了包含1000行数据的测试文件")
            
            # 1. 获取文件信息
            info = service.get_sheet_info(csv_path)
            print(f"文件大小: {info['file_size']} bytes")
            print(f"数据维度: {info['dimensions']['rows']}行 x {info['dimensions']['cols']}列")
            
            assert info['dimensions']['rows'] == 1001  # 包括表头
            assert info['dimensions']['cols'] == 6
            
            # 2. 解析大文件
            json_data = service.parse_sheet(csv_path)
            print(f"处理模式: {json_data['size_info']['processing_mode']}")
            print(f"数据行数: {len(json_data['rows'])}")
            
            # 大文件应该触发特殊处理
            assert json_data['size_info']['processing_mode'] in ['full_with_warning', 'summary']
            
            # 3. 转换为HTML（可能会有大小限制）
            with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as html_file:
                html_path = html_file.name
            
            try:
                html_result = service.convert_to_html(csv_path, html_path)
                print(f"HTML转换状态: {html_result['status']}")
                print(f"HTML文件大小: {html_result['file_size']} bytes")
                
                # 验证大文件处理
                assert html_result['status'] == 'success'
                assert os.path.exists(html_path)
                
                # 检查HTML文件大小是否合理
                html_size = os.path.getsize(html_path)
                print(f"实际HTML文件大小: {html_size} bytes")
                
                # 大文件的HTML应该被适当处理
                assert html_size > 0
                
            finally:
                if os.path.exists(html_path):
                    os.unlink(html_path)
            
            print(f"✅ 大文件处理测试完成")
            
        finally:
            os.unlink(csv_path)
