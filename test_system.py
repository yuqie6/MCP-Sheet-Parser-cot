#!/usr/bin/env python3
"""
系统完整性测试脚本
"""

def test_imports():
    """测试所有关键模块的导入"""
    print("=== 导入测试 ===")
    
    try:
        from src.parsers.factory import ParserFactory
        print("✅ ParserFactory导入成功")
    except Exception as e:
        print(f"❌ ParserFactory导入失败: {e}")
        return False
    
    try:
        from src.models.table_model import Sheet, Row, Cell, Style
        print("✅ 数据模型导入成功")
    except Exception as e:
        print(f"❌ 数据模型导入失败: {e}")
        return False
    
    try:
        from src.converters.html_converter import HTMLConverter
        print("✅ HTMLConverter导入成功")
    except Exception as e:
        print(f"❌ HTMLConverter导入失败: {e}")
        return False
    
    try:
        from src.core_service import CoreService
        print("✅ CoreService导入成功")
    except Exception as e:
        print(f"❌ CoreService导入失败: {e}")
        return False
    
    try:
        from src.mcp_server.tools import register_tools
        print("✅ MCP工具导入成功")
    except Exception as e:
        print(f"❌ MCP工具导入失败: {e}")
        return False
    
    return True

def test_core_functionality():
    """测试核心功能"""
    print("\n=== 核心功能测试 ===")
    
    try:
        from src.parsers.factory import ParserFactory
        factory = ParserFactory()
        formats = factory.get_supported_formats()
        print(f"✅ 支持的格式: {formats}")
    except Exception as e:
        print(f"❌ ParserFactory测试失败: {e}")
        return False
    
    try:
        from src.core_service import CoreService
        service = CoreService()
        
        # 检查方法存在性
        methods = ['convert_to_html', 'parse_sheet', 'apply_changes']
        for method in methods:
            if hasattr(service, method):
                print(f"✅ {method}方法存在")
            else:
                print(f"❌ {method}方法缺失")
                return False
    except Exception as e:
        print(f"❌ CoreService测试失败: {e}")
        return False
    
    try:
        from src.converters.html_converter import HTMLConverter
        converter = HTMLConverter()
        
        if hasattr(converter, 'convert_to_file'):
            print("✅ HTMLConverter.convert_to_file方法存在")
        else:
            print("❌ HTMLConverter.convert_to_file方法缺失")
            return False
    except Exception as e:
        print(f"❌ HTMLConverter测试失败: {e}")
        return False
    
    return True

def test_mcp_tools():
    """测试MCP工具定义"""
    print("\n=== MCP工具测试 ===")
    
    try:
        from mcp.server import Server
        from src.mcp_server.tools import register_tools
        
        server = Server("test-server")
        register_tools(server)
        print("✅ MCP工具注册成功")
        
        return True
    except Exception as e:
        print(f"❌ MCP工具测试失败: {e}")
        return False

def main():
    """主测试函数"""
    print("🔍 MCP Sheet Parser 系统完整性测试")
    print("=" * 50)
    
    success = True
    
    # 测试导入
    if not test_imports():
        success = False
    
    # 测试核心功能
    if not test_core_functionality():
        success = False
    
    # 测试MCP工具
    if not test_mcp_tools():
        success = False
    
    print("\n" + "=" * 50)
    if success:
        print("🎉 所有测试通过！系统状态良好")
    else:
        print("⚠️ 发现问题，需要修复")
    
    return success

if __name__ == "__main__":
    main()
