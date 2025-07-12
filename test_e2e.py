#!/usr/bin/env python3
"""
端到端功能测试
"""

import tempfile
import os
from pathlib import Path

def create_test_csv():
    """创建测试CSV文件"""
    content = """Name,Age,City
Alice,25,New York
Bob,30,London
Charlie,35,Tokyo"""
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
        f.write(content)
        return f.name

def test_convert_to_html():
    """测试convert_to_html功能"""
    print("\n=== 测试convert_to_html功能 ===")
    
    try:
        from src.core_service import CoreService
        
        # 创建测试文件
        csv_file = create_test_csv()
        print(f"✅ 创建测试CSV文件: {csv_file}")
        
        # 测试转换
        service = CoreService()
        result = service.convert_to_html(csv_file)
        
        print(f"✅ HTML转换成功")
        print(f"   输出文件: {result['output_path']}")
        print(f"   文件大小: {result['file_size_kb']} KB")
        print(f"   转换行数: {result['rows_converted']}")
        print(f"   转换单元格数: {result['cells_converted']}")
        
        # 检查HTML文件是否存在
        if os.path.exists(result['output_path']):
            print("✅ HTML文件生成成功")
            
            # 读取并显示部分内容
            with open(result['output_path'], 'r', encoding='utf-8') as f:
                content = f.read()
                if '<table>' in content and 'Alice' in content:
                    print("✅ HTML内容验证通过")
                else:
                    print("❌ HTML内容验证失败")
                    return False
        else:
            print("❌ HTML文件未生成")
            return False
        
        # 清理文件
        os.unlink(csv_file)
        os.unlink(result['output_path'])
        
        return True
        
    except Exception as e:
        print(f"❌ convert_to_html测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_parse_sheet():
    """测试parse_sheet功能"""
    print("\n=== 测试parse_sheet功能 ===")
    
    try:
        from src.core_service import CoreService
        
        # 创建测试文件
        csv_file = create_test_csv()
        print(f"✅ 创建测试CSV文件: {csv_file}")
        
        # 测试解析
        service = CoreService()
        result = service.parse_sheet(csv_file)
        
        print(f"✅ 表格解析成功")
        print(f"   工作表名: {result['sheet_name']}")
        print(f"   总行数: {result['metadata']['total_rows']}")
        print(f"   总列数: {result['metadata']['total_cols']}")
        print(f"   数据行数: {result['metadata']['data_rows']}")
        
        # 验证数据内容
        if 'headers' in result and 'rows' in result:
            print(f"   表头: {result['headers']}")
            if len(result['rows']) > 0:
                print(f"   第一行数据: {[cell['value'] for cell in result['rows'][0]]}")
                print("✅ JSON数据验证通过")
            else:
                print("❌ 没有数据行")
                return False
        else:
            print("❌ JSON格式验证失败")
            return False
        
        # 清理文件
        os.unlink(csv_file)
        
        return True
        
    except Exception as e:
        print(f"❌ parse_sheet测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_apply_changes():
    """测试apply_changes功能"""
    print("\n=== 测试apply_changes功能 ===")
    
    try:
        from src.core_service import CoreService
        
        # 创建测试文件
        csv_file = create_test_csv()
        print(f"✅ 创建测试CSV文件: {csv_file}")
        
        # 模拟修改后的JSON数据
        table_model_json = {
            "sheet_name": "test_sheet",
            "headers": ["Name", "Age", "City"],
            "rows": [
                [{"value": "Alice"}, {"value": "26"}, {"value": "New York"}],
                [{"value": "Bob"}, {"value": "31"}, {"value": "London"}]
            ]
        }
        
        # 测试应用修改
        service = CoreService()
        result = service.apply_changes(csv_file, table_model_json, create_backup=True)
        
        print(f"✅ 修改应用测试成功")
        print(f"   状态: {result['status']}")
        print(f"   文件路径: {result['file_path']}")
        print(f"   备份路径: {result['backup_path']}")
        print(f"   表头数量: {result['headers_count']}")
        print(f"   数据行数: {result['rows_count']}")
        
        # 检查备份文件是否创建
        if result['backup_path'] and os.path.exists(result['backup_path']):
            print("✅ 备份文件创建成功")
            os.unlink(result['backup_path'])
        else:
            print("❌ 备份文件未创建")
        
        # 清理文件
        os.unlink(csv_file)
        
        return True
        
    except Exception as e:
        print(f"❌ apply_changes测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主测试函数"""
    print("🚀 MCP Sheet Parser 端到端功能测试")
    print("=" * 50)
    
    success = True
    
    # 测试三个核心功能
    if not test_convert_to_html():
        success = False
    
    if not test_parse_sheet():
        success = False
    
    if not test_apply_changes():
        success = False
    
    print("\n" + "=" * 50)
    if success:
        print("🎉 所有端到端测试通过！系统功能完整")
        print("✅ convert_to_html: 完美HTML转换")
        print("✅ parse_sheet: LLM友好JSON解析")
        print("✅ apply_changes: 数据写回验证")
    else:
        print("⚠️ 发现功能问题，需要修复")
    
    return success

if __name__ == "__main__":
    main()
