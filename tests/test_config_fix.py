#!/usr/bin/env python3
"""
测试配置系统修复

验证统一配置系统是否正常工作，确保没有配置冲突。
"""

import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_unified_config():
    """测试统一配置系统"""
    print("🔍 测试统一配置系统")
    print("-" * 40)
    
    try:
        from src.unified_config import get_config, get_cache_config, get_config_manager
        
        # 测试基本配置获取
        config = get_config()
        print(f"✅ 基本配置获取成功")
        print(f"   - 最大文件大小: {config.max_file_size_mb}MB")
        print(f"   - 缓存启用: {config.cache_enabled}")
        print(f"   - 小文件阈值: {config.small_file_threshold_cells} 单元格")
        
        # 测试缓存配置获取
        cache_config = get_cache_config()
        print(f"✅ 缓存配置获取成功")
        print(f"   - 缓存启用: {cache_config.cache_enabled}")
        print(f"   - 最大条目数: {cache_config.max_entries}")
        print(f"   - 磁盘缓存启用: {cache_config.disk_cache_enabled}")
        print(f"   - 内存缓存启用: {cache_config.memory_cache_enabled}")
        
        # 测试配置管理器
        config_manager = get_config_manager()
        print(f"✅ 配置管理器获取成功")
        
        # 测试配置验证
        config.validate()
        cache_config.validate()
        print(f"✅ 配置验证通过")
        
        assert True
        
    except Exception as e:
        print(f"❌ 统一配置系统测试失败: {e}")
        import traceback
        traceback.print_exc()
        assert False, f"统一配置系统测试失败: {e}"

def test_cache_manager():
    """测试缓存管理器"""
    print("\n🔍 测试缓存管理器")
    print("-" * 40)
    
    try:
        from src.cache import get_cache_manager
        
        # 获取缓存管理器
        cache_manager = get_cache_manager()
        print(f"✅ 缓存管理器创建成功")
        
        # 测试缓存配置
        config = cache_manager.config
        print(f"   - 缓存启用: {config.cache_enabled}")
        print(f"   - 最大条目数: {config.max_entries}")
        print(f"   - 缓存过期时间: {config.cache_expiry_seconds}秒")
        
        # 测试缓存操作
        test_key = "test_file.xlsx"
        test_data = {"test": "data"}
        
        # 设置缓存
        cache_manager.set(test_key, test_data)
        print(f"✅ 缓存设置成功")
        
        # 获取缓存
        cached_data = cache_manager.get(test_key)
        if cached_data and cached_data.get('data') == test_data:
            print(f"✅ 缓存获取成功")
        else:
            print(f"⚠️  缓存获取结果不匹配")
        assert cached_data and cached_data.get('data') == test_data, "缓存获取结果不匹配"
        
        # 获取缓存统计
        stats = cache_manager.get_stats()
        print(f"✅ 缓存统计获取成功")
        print(f"   - 配置信息: {len(stats.get('config', {}))} 项")
        
        assert True
        
    except Exception as e:
        print(f"❌ 缓存管理器测试失败: {e}")
        import traceback
        traceback.print_exc()
        assert False, f"缓存管理器测试失败: {e}"

def test_config_imports():
    """测试配置导入是否正常"""
    print("\n🔍 测试配置导入")
    print("-" * 40)
    
    try:
        # 测试核心服务的配置导入
        from src.core_service import CoreService
        core_service = CoreService()
        print(f"✅ 核心服务配置导入成功")
        
        # 测试验证器的配置导入
        from src.validators import FileValidator
        max_size = FileValidator.get_max_file_size_mb()
        print(f"✅ 验证器配置导入成功 (最大文件大小: {max_size}MB)")
        
        # 测试字体管理器的配置导入
        from src.font_manager import FontManager
        font_manager = FontManager()
        print(f"✅ 字体管理器配置导入成功")
        
        assert True
        
    except Exception as e:
        print(f"❌ 配置导入测试失败: {e}")
        import traceback
        traceback.print_exc()
        assert False, f"配置导入测试失败: {e}"

def test_no_config_conflicts():
    """测试是否存在配置冲突"""
    print("\n🔍 测试配置冲突")
    print("-" * 40)
    
    try:
        # 检查是否还有旧的配置导入
        import importlib.util
        
        # 检查是否存在 src/cache/config.py
        cache_config_path = project_root / "src" / "cache" / "config.py"
        assert not cache_config_path.exists(), f"发现旧的配置文件: {cache_config_path}"
        print(f"✅ 没有发现旧的配置文件")
        
        # 检查是否有其他config.py文件（除了font_config.json）
        config_files = []
        for config_file in project_root.rglob("config.py"):
            if "site-packages" not in str(config_file) and ".venv" not in str(config_file):
                config_files.append(config_file)
        
        if config_files:
            print(f"⚠️  发现其他配置文件: {config_files}")
            for file in config_files:
                print(f"     - {file}")
        else:
            print(f"✅ 没有发现冲突的配置文件")
        
        # 测试统一配置是否正常工作
        from src.unified_config import get_config
        config1 = get_config()
        config2 = get_config()
        
        # 应该是同一个实例（单例模式）
        assert config1 is config2, "配置单例模式可能有问题"
        print(f"✅ 配置单例模式工作正常")
        
        assert len(config_files) == 0, f"发现其他配置文件: {config_files}"
        
    except Exception as e:
        print(f"❌ 配置冲突测试失败: {e}")
        import traceback
        traceback.print_exc()
        assert False, f"配置冲突测试失败: {e}"
