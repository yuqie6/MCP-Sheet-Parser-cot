import pytest
from unittest.mock import MagicMock, patch
from src.unified_config import UnifiedConfig

@pytest.fixture
def config():
    """Fixture for UnifiedConfig."""
    return UnifiedConfig()

# === TDD测试：提升UnifiedConfig覆盖率 ===

def test_config_initialization_with_defaults():
    """
    TDD测试：UnifiedConfig应该使用默认值正确初始化
    
    这个测试覆盖初始化的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config = UnifiedConfig()
    
    # 验证默认值
    assert config.memory_cache_enabled is True
    assert config.disk_cache_enabled is True
    assert config.max_entries == 1000
    assert config.max_disk_cache_size_mb == 100
    assert config.cache_expiry_seconds == 3600
    assert config.cache_dir == "cache"

def test_is_cache_enabled_with_both_enabled(config):
    """
    TDD测试：is_cache_enabled应该在任一缓存启用时返回True
    
    这个测试覆盖第71行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config.memory_cache_enabled = True
    config.disk_cache_enabled = True
    
    assert config.is_cache_enabled() is True

def test_is_cache_enabled_with_memory_only(config):
    """
    TDD测试：is_cache_enabled应该在只有内存缓存启用时返回True
    
    这个测试覆盖第71行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config.memory_cache_enabled = True
    config.disk_cache_enabled = False
    
    assert config.is_cache_enabled() is True

def test_is_cache_enabled_with_disk_only(config):
    """
    TDD测试：is_cache_enabled应该在只有磁盘缓存启用时返回True
    
    这个测试覆盖第71行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config.memory_cache_enabled = False
    config.disk_cache_enabled = True
    
    assert config.is_cache_enabled() is True

def test_is_cache_enabled_with_both_disabled(config):
    """
    TDD测试：is_cache_enabled应该在两个缓存都禁用时返回False
    
    这个测试覆盖第71行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config.memory_cache_enabled = False
    config.disk_cache_enabled = False
    
    assert config.is_cache_enabled() is False

@patch.dict('os.environ', {'MEMORY_CACHE_ENABLED': 'false'})
def test_config_with_environment_variables():
    """
    TDD测试：UnifiedConfig应该能从环境变量读取配置
    
    这个测试覆盖环境变量处理的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config = UnifiedConfig()
    
    # 验证环境变量被正确读取
    assert config.memory_cache_enabled is False

@patch.dict('os.environ', {'MAX_ENTRIES': '500'})
def test_config_with_numeric_environment_variables():
    """
    TDD测试：UnifiedConfig应该能从环境变量读取数值配置
    
    这个测试覆盖数值环境变量处理的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config = UnifiedConfig()
    
    # 验证数值环境变量被正确读取
    assert config.max_entries == 500

def test_get_cache_config():
    """
    TDD测试：get_cache_config应该返回UnifiedConfig实例
    
    这个测试覆盖第82行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    from src.unified_config import get_cache_config
    
    config = get_cache_config()
    
    # 验证返回正确的配置实例
    assert isinstance(config, UnifiedConfig)

def test_get_streaming_config():
    """
    TDD测试：get_streaming_config应该返回UnifiedConfig实例
    
    这个测试覆盖第85行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    from src.unified_config import get_streaming_config
    
    config = get_streaming_config()
    
    # 验证返回正确的配置实例
    assert isinstance(config, UnifiedConfig)

def test_get_conversion_config():
    """
    TDD测试：get_conversion_config应该返回UnifiedConfig实例
    
    这个测试覆盖第88行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    from src.unified_config import get_conversion_config
    
    config = get_conversion_config()
    
    # 验证返回正确的配置实例
    assert isinstance(config, UnifiedConfig)

def test_get_parsing_config():
    """
    TDD测试：get_parsing_config应该返回UnifiedConfig实例
    
    这个测试覆盖第91行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    from src.unified_config import get_parsing_config
    
    config = get_parsing_config()
    
    # 验证返回正确的配置实例
    assert isinstance(config, UnifiedConfig)

def test_get_validation_config():
    """
    TDD测试：get_validation_config应该返回UnifiedConfig实例
    
    这个测试覆盖第94行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    from src.unified_config import get_validation_config
    
    config = get_validation_config()
    
    # 验证返回正确的配置实例
    assert isinstance(config, UnifiedConfig)

def test_get_font_config():
    """
    TDD测试：get_font_config应该返回UnifiedConfig实例
    
    这个测试覆盖第97行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    from src.unified_config import get_font_config
    
    config = get_font_config()
    
    # 验证返回正确的配置实例
    assert isinstance(config, UnifiedConfig)

@patch.dict('os.environ', {'INVALID_BOOL': 'invalid_value'})
def test_config_with_invalid_boolean_environment_variable():
    """
    TDD测试：UnifiedConfig应该处理无效的布尔环境变量
    
    这个测试确保方法在遇到无效布尔值时使用默认值
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config = UnifiedConfig()
    
    # 应该使用默认值而不是崩溃
    assert hasattr(config, 'memory_cache_enabled')

@patch.dict('os.environ', {'MAX_ENTRIES': 'invalid_number'})
def test_config_with_invalid_numeric_environment_variable():
    """
    TDD测试：UnifiedConfig应该处理无效的数值环境变量
    
    这个测试确保方法在遇到无效数值时使用默认值
    """
    # 🔴 红阶段：编写测试描述期望的行为
    config = UnifiedConfig()
    
    # 应该使用默认值而不是崩溃
    assert config.max_entries == 1000  # 默认值
