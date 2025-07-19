
import pytest
from unittest.mock import MagicMock, patch, mock_open
import time
from src.cache.cache_manager import CacheManager, get_cache_manager, reset_cache_manager
from src.unified_config import UnifiedConfig

@pytest.fixture
def mock_config():
    """Fixture for a mocked cache config."""
    config = MagicMock(spec=UnifiedConfig)
    config.is_cache_enabled.return_value = True
    config.memory_cache_enabled = True
    config.disk_cache_enabled = True
    config.max_entries = 10
    config.max_disk_cache_size_mb = 10
    config.cache_expiry_seconds = 3600
    config.cache_dir = "/tmp/cache"
    return config

@pytest.fixture
def cache_manager(mock_config):
    """Fixture for CacheManager with mocked dependencies."""
    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
         patch('src.cache.cache_manager.DiskCache') as mock_disk:
        manager = CacheManager(config=mock_config)
        manager.memory_cache = mock_lru.return_value
        manager.disk_cache = mock_disk.return_value
    return manager

@patch('src.cache.cache_manager.Path.exists', return_value=True)
@patch('src.cache.cache_manager.Path.stat')
def test_cache_manager_set_get(mock_stat, mock_exists, cache_manager):
    """Test setting and getting a cache item."""
    mock_stat.return_value.st_size = 100
    mock_stat.return_value.st_mtime_ns = 12345
    cache_manager.memory_cache.get.return_value = None
    cache_manager.disk_cache.get.return_value = None

    cache_manager.set("file.xlsx", {"data": "value"})
    cache_manager.memory_cache.set.assert_called_once()
    cache_manager.disk_cache.set.assert_called_once()

    cache_manager.get("file.xlsx")
    cache_manager.memory_cache.get.assert_called()

def test_get_cache_manager_singleton():
    """Test that get_cache_manager returns a singleton."""
    # Reset for a clean test
    reset_cache_manager()
    manager1 = get_cache_manager()
    manager2 = get_cache_manager()
    assert manager1 is manager2

def test_reset_cache_manager():
    """Test resetting the cache manager."""
    reset_cache_manager()
    manager1 = get_cache_manager()
    reset_cache_manager()
    manager2 = get_cache_manager()
    assert manager1 is not manager2

def test_is_cache_valid(cache_manager):
    """Test cache validity check."""
    valid_entry = {'timestamp': time.time(), 'file_path': 'file.xlsx'}
    expired_entry = {'timestamp': time.time() - 4000, 'file_path': 'file.xlsx'}
    with patch('src.cache.cache_manager.Path.exists', return_value=True):
        assert cache_manager._is_cache_valid(valid_entry) is True
        assert cache_manager._is_cache_valid(expired_entry) is False

def test_clear_cache(cache_manager):
    """Test clearing the cache."""
    cache_manager.clear()
    cache_manager.memory_cache.clear.assert_called_once()
    cache_manager.disk_cache.clear.assert_called_once()

def test_get_stats(cache_manager):
    """Test getting cache stats."""
    cache_manager.memory_cache.cache.hits = 10
    cache_manager.memory_cache.cache.misses = 5
    cache_manager.memory_cache.cache.maxsize = 100
    cache_manager.memory_cache.cache.__len__.return_value = 20
    stats = cache_manager.get_stats()
    assert stats['memory_cache']['hits'] == 10
    assert stats['memory_cache']['misses'] == 5

# === TDD测试：提升CacheManager覆盖率 ===

def test_cache_manager_initialization_with_disabled_cache():
    """
    TDD测试：CacheManager应该在缓存禁用时正确初始化

    这个测试覆盖缓存禁用时的初始化代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = False
    mock_config.memory_cache_enabled = False
    mock_config.disk_cache_enabled = False

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config):
        manager = CacheManager(config=mock_config)

        # 验证缓存被禁用
        assert manager.memory_cache is None
        assert manager.disk_cache is None

def test_cache_manager_initialization_memory_only():
    """
    TDD测试：CacheManager应该支持仅内存缓存模式

    这个测试覆盖仅启用内存缓存的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = True
    mock_config.memory_cache_enabled = True
    mock_config.disk_cache_enabled = False
    mock_config.max_entries = 10
    mock_config.cache_expiry_seconds = 3600

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru:

        manager = CacheManager(config=mock_config)

        # 验证只有内存缓存被初始化
        assert manager.memory_cache is not None
        assert manager.disk_cache is None
        mock_lru.assert_called_once()

def test_cache_manager_initialization_disk_only():
    """
    TDD测试：CacheManager应该支持仅磁盘缓存模式

    这个测试覆盖仅启用磁盘缓存的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = True
    mock_config.memory_cache_enabled = False
    mock_config.disk_cache_enabled = True
    mock_config.max_disk_cache_size_mb = 10
    mock_config.cache_dir = "/tmp/cache"

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('src.cache.cache_manager.DiskCache') as mock_disk:

        manager = CacheManager(config=mock_config)

        # 验证只有磁盘缓存被初始化
        assert manager.memory_cache is None
        assert manager.disk_cache is not None
        mock_disk.assert_called_once()

def test_cache_manager_get_with_disabled_cache():
    """
    TDD测试：CacheManager在缓存禁用时get应该返回None

    这个测试确保禁用缓存时get方法的行为
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = False

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config):
        manager = CacheManager(config=mock_config)

        result = manager.get("test_file.xlsx")
        assert result is None

def test_cache_manager_set_with_disabled_cache():
    """
    TDD测试：CacheManager在缓存禁用时set应该不执行任何操作

    这个测试确保禁用缓存时set方法的行为
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = False

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config):
        manager = CacheManager(config=mock_config)

        # 这应该不会抛出异常，即使缓存被禁用
        manager.set("test_file.xlsx", {"data": "test"})

        # 验证没有缓存实例
        assert manager.memory_cache is None
        assert manager.disk_cache is None

def test_cache_manager_generate_cache_key_with_all_parameters():
    """
    TDD测试：generate_cache_key应该能处理所有参数

    这个测试覆盖第84-94行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = True

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch.object(CacheManager, '_calculate_file_hash', return_value='abc123'):

        manager = CacheManager(config=mock_config)

        # 测试包含所有参数的缓存键生成
        cache_key = manager.generate_cache_key(
            file_path="test.xlsx",
            range_string="A1:B10",
            sheet_name="Sheet1"
        )

        # 验证缓存键格式
        expected_key = "abc123|range:A1:B10|sheet:Sheet1"
        assert cache_key == expected_key

def test_cache_manager_generate_cache_key_with_partial_parameters():
    """
    TDD测试：generate_cache_key应该能处理部分参数

    这个测试覆盖第88-91行的条件分支
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = True

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch.object(CacheManager, '_calculate_file_hash', return_value='def456'):

        manager = CacheManager(config=mock_config)

        # 测试只有range_string的情况
        cache_key = manager.generate_cache_key(
            file_path="test.xlsx",
            range_string="A1:B10"
        )
        expected_key = "def456|range:A1:B10"
        assert cache_key == expected_key

        # 测试只有sheet_name的情况
        cache_key = manager.generate_cache_key(
            file_path="test.xlsx",
            sheet_name="Sheet1"
        )
        expected_key = "def456|sheet:Sheet1"
        assert cache_key == expected_key

        # 测试只有file_path的情况
        cache_key = manager.generate_cache_key(file_path="test.xlsx")
        expected_key = "def456"
        assert cache_key == expected_key

@patch('src.cache.cache_manager.Path.exists', return_value=True)
@patch('src.cache.cache_manager.Path.stat')
def test_cache_manager_calculate_file_hash_small_file(mock_stat, mock_exists):
    """
    TDD测试：_calculate_file_hash应该为小文件计算内容哈希

    这个测试覆盖小文件的哈希计算代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = True

    # 模拟小文件（小于1MB）
    mock_stat.return_value.st_size = 500000  # 500KB
    mock_stat.return_value.st_mtime_ns = 1234567890

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('builtins.open', mock_open(read_data=b'test file content')), \
         patch('hashlib.sha256') as mock_sha256:

        mock_hash = MagicMock()
        mock_hash.hexdigest.return_value = 'content_hash_123'
        mock_sha256.return_value = mock_hash

        manager = CacheManager(config=mock_config)
        file_hash = manager._calculate_file_hash("small_file.xlsx")

        # 验证包含了内容哈希
        assert 'content_hash_123' in file_hash

@patch('src.cache.cache_manager.Path.exists', return_value=True)
@patch('src.cache.cache_manager.Path.stat')
def test_cache_manager_calculate_file_hash_large_file(mock_stat, mock_exists):
    """
    TDD测试：_calculate_file_hash应该为大文件只计算元数据哈希

    这个测试覆盖大文件的哈希计算代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = True

    # 模拟大文件（大于1MB）
    mock_stat.return_value.st_size = 2000000  # 2MB
    mock_stat.return_value.st_mtime_ns = 1234567890

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('hashlib.sha256') as mock_sha256:

        mock_hash = MagicMock()
        mock_hash.hexdigest.return_value = 'metadata_hash_456'
        mock_sha256.return_value = mock_hash

        manager = CacheManager(config=mock_config)
        file_hash = manager._calculate_file_hash("large_file.xlsx")

        # 验证返回了哈希值
        assert file_hash == 'metadata_hash_456'

@patch('src.cache.cache_manager.Path.exists', return_value=False)
def test_cache_manager_calculate_file_hash_nonexistent_file(mock_exists):
    """
    TDD测试：_calculate_file_hash应该处理不存在的文件

    这个测试确保方法在文件不存在时的行为
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = True

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config):
        manager = CacheManager(config=mock_config)

        # 应该返回基于文件路径的哈希
        file_hash = manager._calculate_file_hash("nonexistent.xlsx")

        # 验证返回了某种哈希值
        assert isinstance(file_hash, str)
        assert len(file_hash) > 0
