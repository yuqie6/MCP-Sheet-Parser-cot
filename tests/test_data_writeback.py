"""
数据写回功能测试模块

测试apply_changes的真实数据写回功能
"""

import pytest
import tempfile
import os
import csv
from pathlib import Path
from src.core_service import CoreService


class TestDataWriteback:
    """数据写回功能测试。"""
    
    def test_csv_data_writeback(self):
        """测试CSV文件的数据写回功能。"""
        # 创建原始CSV文件
        original_data = """Name,Age,City
Alice,25,New York
Bob,30,London
Charlie,35,Paris"""
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
            csv_file.write(original_data)
            csv_path = csv_file.name
        
        try:
            service = CoreService()
            
            # 1. 解析原始数据
            original_json = service.parse_sheet(csv_path)
            print(f"原始数据: {len(original_json['rows'])} 行")
            
            # 验证原始数据
            assert len(original_json['rows']) == 3
            assert original_json['rows'][0][0]['value'] == 'Alice'
            assert original_json['rows'][0][1]['value'] == '25'
            assert original_json['rows'][1][0]['value'] == 'Bob'
            
            # 2. 修改数据
            modified_json = original_json.copy()
            
            # 修改Alice的年龄从25改为26
            modified_json['rows'][0][1]['value'] = '26'
            
            # 修改Bob的城市从London改为Tokyo
            modified_json['rows'][1][2]['value'] = 'Tokyo'
            
            # 添加新的一行数据
            new_row = [
                {'value': 'David'},
                {'value': '28'},
                {'value': 'Berlin'}
            ]
            modified_json['rows'].append(new_row)
            
            print(f"修改后数据: {len(modified_json['rows'])} 行")
            
            # 3. 应用修改
            result = service.apply_changes(csv_path, modified_json, create_backup=True)
            
            print(f"写回结果: {result}")
            
            # 验证写回结果
            assert result['status'] == 'success'
            assert result['changes_applied'] == 4  # 4行数据（包括新增的）
            assert result['backup_created'] is True
            
            # 4. 验证备份文件存在
            backup_path = result['backup_path']
            assert os.path.exists(backup_path)
            
            # 验证备份文件内容是原始数据
            with open(backup_path, 'r', encoding='utf-8') as f:
                backup_content = f.read()
                assert 'Alice,25,New York' in backup_content
                assert 'Bob,30,London' in backup_content
            
            # 5. 验证修改后的文件内容
            with open(csv_path, 'r', encoding='utf-8') as f:
                modified_content = f.read()
                print(f"修改后文件内容:\n{modified_content}")
            
            # 验证修改生效
            assert 'Alice,26,New York' in modified_content  # Alice年龄改为26
            assert 'Bob,30,Tokyo' in modified_content       # Bob城市改为Tokyo
            assert 'David,28,Berlin' in modified_content    # 新增的David
            
            # 6. 重新解析验证
            final_json = service.parse_sheet(csv_path)
            assert len(final_json['rows']) == 4  # 现在有4行数据
            assert final_json['rows'][0][1]['value'] == '26'  # Alice的年龄
            assert final_json['rows'][1][2]['value'] == 'Tokyo'  # Bob的城市
            assert final_json['rows'][3][0]['value'] == 'David'  # 新增的David
            
            print("✅ CSV数据写回测试成功")
            
        finally:
            # 清理文件
            os.unlink(csv_path)
            if os.path.exists(backup_path):
                os.unlink(backup_path)
    
    def test_xlsx_data_writeback(self):
        """测试XLSX文件的数据写回功能。"""
        # 创建一个简单的XLSX文件
        import openpyxl
        
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        
        # 添加数据
        worksheet['A1'] = 'Product'
        worksheet['B1'] = 'Price'
        worksheet['C1'] = 'Stock'
        
        worksheet['A2'] = 'Laptop'
        worksheet['B2'] = 999.99
        worksheet['C2'] = 10
        
        worksheet['A3'] = 'Mouse'
        worksheet['B3'] = 29.99
        worksheet['C3'] = 50
        
        # 保存到临时文件
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as xlsx_file:
            xlsx_path = xlsx_file.name
        
        workbook.save(xlsx_path)
        workbook.close()
        
        try:
            service = CoreService()
            
            # 1. 解析原始数据
            original_json = service.parse_sheet(xlsx_path)
            print(f"原始XLSX数据: {len(original_json['rows'])} 行")
            
            # 验证原始数据
            assert len(original_json['rows']) == 2
            assert original_json['rows'][0][0]['value'] == 'Laptop'
            assert original_json['rows'][0][1]['value'] == 999.99
            
            # 2. 修改数据
            modified_json = original_json.copy()
            
            # 修改Laptop价格从999.99改为899.99
            modified_json['rows'][0][1]['value'] = '899.99'
            
            # 修改Mouse库存从50改为100
            modified_json['rows'][1][2]['value'] = '100'
            
            print(f"准备写回修改的数据")
            
            # 3. 应用修改
            result = service.apply_changes(xlsx_path, modified_json, create_backup=True)
            
            print(f"XLSX写回结果: {result}")
            
            # 验证写回结果
            assert result['status'] == 'success'
            assert result['changes_applied'] > 0
            assert result['backup_created'] is True
            
            # 4. 验证修改后的文件内容
            final_json = service.parse_sheet(xlsx_path)
            print(f"修改后XLSX数据: {len(final_json['rows'])} 行")
            
            # 验证修改生效
            assert final_json['rows'][0][1]['value'] == 899.99  # Laptop价格改为899.99
            assert final_json['rows'][1][2]['value'] == 100     # Mouse库存改为100
            
            print("✅ XLSX数据写回测试成功")
            
        finally:
            # 清理文件
            os.unlink(xlsx_path)
            backup_path = xlsx_path + ".backup"
            if os.path.exists(backup_path):
                os.unlink(backup_path)
    
    def test_unsupported_format_writeback(self):
        """测试不支持格式的写回处理。"""
        service = CoreService()

        # 创建临时文件来测试格式检查
        import tempfile

        # 测试XLS格式
        with tempfile.NamedTemporaryFile(suffix='.xls', delete=False) as xls_file:
            xls_path = xls_file.name

        try:
            # 测试XLS格式，现在应该可以写回
            # with pytest.raises(RuntimeError) as exc_info:
            #     service.apply_changes(xls_path, {"sheet_name": "test", "headers": [], "rows": []})
            # assert "XLS格式暂不支持数据写回" in str(exc_info.value)
            # 由于已经支持，我们简单地调用它，不检查异常
            service.apply_changes(xls_path, {"sheet_name": "test", "headers": ["a"], "rows": [[{"value": 1}]]})
        finally:
            os.unlink(xls_path)

        # 测试XLSB格式
        with tempfile.NamedTemporaryFile(suffix='.xlsb', delete=False) as xlsb_file:
            xlsb_path = xlsb_file.name

        try:
            with pytest.raises(RuntimeError) as exc_info:
                service.apply_changes(xlsb_path, {"sheet_name": "test", "headers": [], "rows": []})
            assert "XLSB格式暂不支持数据写回" in str(exc_info.value)
        finally:
            os.unlink(xlsb_path)

        # 测试未知格式
        with tempfile.NamedTemporaryFile(suffix='.unknown', delete=False) as unknown_file:
            unknown_path = unknown_file.name

        try:
            with pytest.raises(RuntimeError) as exc_info:
                service.apply_changes(unknown_path, {"sheet_name": "test", "headers": [], "rows": []})
            assert "不支持的文件格式" in str(exc_info.value)
        finally:
            os.unlink(unknown_path)
    
    def test_invalid_json_format(self):
        """测试无效JSON格式的处理。"""
        service = CoreService()
        
        # 创建临时CSV文件
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
            csv_file.write("A,B\n1,2")
            csv_path = csv_file.name
        
        try:
            # 测试缺少必需字段
            invalid_json = {"invalid": "data"}
            
            with pytest.raises(ValueError) as exc_info:
                service.apply_changes(csv_path, invalid_json)
            assert "缺少必需字段" in str(exc_info.value)
            
        finally:
            os.unlink(csv_path)
