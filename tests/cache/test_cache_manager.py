
"""
缓存管理器测试模块

测试缓存管理器的核心功能：初始化、存取、清理、并发安全等。
"""

import pytest
from unittest.mock import MagicMock, patch, mock_open
import time
from src.cache.cache_manager import CacheManager, get_cache_manager, reset_cache_manager
from src.unified_config import UnifiedConfig

def create_mock_config(memory_enabled=True, disk_enabled=True, cache_enabled=True):
    """创建mock配置对象"""
    mock_config = MagicMock(spec=UnifiedConfig)
    mock_config.is_cache_enabled.return_value = cache_enabled
    mock_config.memory_cache_enabled = memory_enabled
    mock_config.disk_cache_enabled = disk_enabled
    mock_config.max_entries = 1000
    mock_config.cache_expiry_seconds = 3600
    mock_config.max_disk_cache_size_mb = 100
    mock_config.cache_dir = "cache"
    return mock_config

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
    mock_config = create_mock_config(memory_enabled=False, disk_enabled=False, cache_enabled=False)

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
    mock_config = create_mock_config(memory_enabled=True, disk_enabled=False)
    mock_config.max_entries = 10  # 覆盖默认值

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
    mock_config = create_mock_config(memory_enabled=False, disk_enabled=True)
    mock_config.max_disk_cache_size_mb = 10  # 覆盖默认值
    mock_config.cache_dir = "/tmp/cache"  # 覆盖默认值

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
    mock_config = create_mock_config(memory_enabled=False, disk_enabled=False, cache_enabled=False)

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config):
        manager = CacheManager(config=mock_config)

        result = manager.get("test_file.xlsx")
        assert result is None

def test_cache_manager_set_with_disabled_cache():
    """
    TDD测试：CacheManager在缓存禁用时set应该不执行任何操作

    这个测试确保禁用缓存时set方法的行为
    """
    mock_config = create_mock_config(memory_enabled=False, disk_enabled=False, cache_enabled=False)

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
    mock_config = create_mock_config()

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch.object(CacheManager, '_calculate_file_hash', return_value='abc123'):

        manager = CacheManager(config=mock_config)

        # 测试包含所有参数的缓存键生成
        cache_key = manager._generate_cache_key(
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
    mock_config = create_mock_config()

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch.object(CacheManager, '_calculate_file_hash', return_value='def456'):

        manager = CacheManager(config=mock_config)

        # 测试只有range_string的情况
        cache_key = manager._generate_cache_key(
            file_path="test.xlsx",
            range_string="A1:B10"
        )
        expected_key = "def456|range:A1:B10"
        assert cache_key == expected_key

        # 测试只有sheet_name的情况
        cache_key = manager._generate_cache_key(
            file_path="test.xlsx",
            sheet_name="Sheet1"
        )
        expected_key = "def456|sheet:Sheet1"
        assert cache_key == expected_key

        # 测试只有file_path的情况
        cache_key = manager._generate_cache_key(file_path="test.xlsx")
        expected_key = "def456"
        assert cache_key == expected_key

def test_cache_manager_calculate_file_hash_small_file(tmp_path):
    """
    TDD测试：_calculate_file_hash应该为小文件计算内容哈希

    这个测试覆盖小文件的哈希计算代码路径
    """
    mock_config = create_mock_config()
    mock_config.cache_dir = str(tmp_path)  # 使用临时目录

    # 创建一个测试文件
    test_file = tmp_path / "test.xlsx"
    test_file.write_bytes(b'test file content')

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('hashlib.sha256') as mock_sha256:

        mock_hash = MagicMock()
        mock_hash.hexdigest.return_value = 'content_hash_123'
        mock_sha256.return_value = mock_hash

        manager = CacheManager(config=mock_config)
        file_hash = manager._calculate_file_hash(str(test_file))

        # 验证包含了内容哈希
        assert 'content_hash_123' in file_hash

def test_cache_manager_calculate_file_hash_large_file(tmp_path):
    """
    TDD测试：_calculate_file_hash应该为大文件只计算元数据哈希

    这个测试覆盖大文件的哈希计算代码路径
    """
    mock_config = create_mock_config()

    # 创建一个真实的大文件用于测试
    large_file = tmp_path / "large_file.xlsx"
    large_file.write_bytes(b"x" * (2 * 1024 * 1024))  # 2MB文件

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('hashlib.sha256') as mock_sha256:

        mock_hash = MagicMock()
        mock_hash.hexdigest.return_value = 'metadata_hash_456'
        mock_sha256.return_value = mock_hash

        manager = CacheManager(config=mock_config)
        file_hash = manager._calculate_file_hash(str(large_file))

        # 验证返回了哈希值
        assert file_hash == 'metadata_hash_456'

@patch('src.cache.cache_manager.Path.exists', return_value=False)
def test_cache_manager_calculate_file_hash_nonexistent_file(mock_exists):
    """
    TDD测试：_calculate_file_hash应该处理不存在的文件

    这个测试确保方法在文件不存在时的行为
    """
    mock_config = create_mock_config()

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config):
        manager = CacheManager(config=mock_config)

        # 应该返回基于文件路径的哈希
        file_hash = manager._calculate_file_hash("nonexistent.xlsx")

        # 验证返回了某种哈希值
        assert isinstance(file_hash, str)
        assert len(file_hash) > 0

class TestCacheManagerInitializationEdgeCases:
    """测试CacheManager初始化的边界情况。"""

    @patch('os.name', 'nt')  # 模拟Windows系统
    def test_cache_manager_initialization_windows_no_cache_dir(self):
        """
        TDD测试：CacheManager应该在Windows上正确处理缺失的cache_dir

        这个测试覆盖第57-58行的Windows缓存目录逻辑
        """
        mock_config = create_mock_config()
        mock_config.cache_dir = None  # 没有设置缓存目录

        with patch.dict('os.environ', {'LOCALAPPDATA': 'C:/Users/Test/AppData/Local'}):
            with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
                 patch('src.cache.cache_manager.LRURowBlockCache'), \
                 patch('src.cache.cache_manager.DiskCache') as mock_disk:

                manager = CacheManager(config=mock_config)

                # 验证DiskCache使用了正确的Windows路径
                mock_disk.assert_called_once()
                call_args = mock_disk.call_args[1]
                assert 'C:/Users/Test/AppData/Local' in call_args['cache_dir'] or \
                       'mcp-sheet-parser' in call_args['cache_dir']

    def test_cache_manager_initialization_unix_no_cache_dir(self):
        """
        TDD测试：CacheManager应该在Unix上正确处理缺失的cache_dir

        这个测试覆盖第59-61行的Unix缓存目录逻辑
        """
        mock_config = create_mock_config()
        mock_config.cache_dir = None  # 没有设置缓存目录

        # 直接模拟Unix逻辑而不是创建实际的Path对象
        with patch('os.name', 'posix'), \
             patch.dict('os.environ', {'XDG_CACHE_HOME': '/home/test/.cache'}), \
             patch('src.cache.cache_manager.Path') as mock_path:

            # 模拟Path操作
            mock_path.return_value.__truediv__.return_value = '/home/test/.cache/mcp-sheet-parser'

            with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
                 patch('src.cache.cache_manager.LRURowBlockCache'), \
                 patch('src.cache.cache_manager.DiskCache') as mock_disk:

                manager = CacheManager(config=mock_config)

                # 验证DiskCache被调用
                mock_disk.assert_called_once()

    def test_cache_manager_initialization_fallback_cache_dir(self):
        """
        TDD测试：CacheManager应该使用后备缓存目录

        这个测试覆盖第64行的后备逻辑
        """
        mock_config = create_mock_config()
        mock_config.cache_dir = None  # 没有设置缓存目录

        # 模拟所有环境变量都不存在的情况
        with patch.dict('os.environ', {}, clear=True):
            with patch('os.path.expanduser', return_value='/home/user'):
                with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
                     patch('src.cache.cache_manager.LRURowBlockCache'), \
                     patch('src.cache.cache_manager.DiskCache') as mock_disk:

                    manager = CacheManager(config=mock_config)

                    # 验证使用了后备路径
                    mock_disk.assert_called_once()
                    call_args = mock_disk.call_args[1]
                    assert '/home/user' in call_args['cache_dir'] or \
                           'mcp-sheet-parser' in call_args['cache_dir']

class TestCacheManagerFileHashExceptions:
    """测试文件哈希计算的异常处理。"""

    def test_calculate_file_hash_general_exception(self):
        """
        TDD测试：_calculate_file_hash应该处理一般异常

        这个测试覆盖第130-133行的异常处理代码
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache'), \
             patch('src.cache.cache_manager.DiskCache'):

            manager = CacheManager(config=mock_config)

            # 模拟Path.exists抛出异常
            with patch('src.cache.cache_manager.Path.exists', side_effect=RuntimeError("Unexpected error")):
                file_hash = manager._calculate_file_hash("test.xlsx")

                # 验证返回了错误标识
                assert isinstance(file_hash, str)
                assert file_hash.startswith("error:")
                assert ":" in file_hash  # 应该包含时间戳

    def test_calculate_file_hash_memory_error(self):
        """
        TDD测试：_calculate_file_hash应该处理内存错误

        这个测试覆盖第125-128行的MemoryError处理代码
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache'), \
             patch('src.cache.cache_manager.DiskCache'):

            manager = CacheManager(config=mock_config)

            # 模拟文件存在但读取时出现内存错误
            with patch('src.cache.cache_manager.Path.exists', return_value=True), \
                 patch('src.cache.cache_manager.Path.stat') as mock_stat, \
                 patch('builtins.open', side_effect=MemoryError("Out of memory")):

                mock_stat.return_value.st_size = 1000000  # 大文件
                mock_stat.return_value.st_mtime_ns = 12345

                file_hash = manager._calculate_file_hash("large_file.xlsx")

                # 验证返回了基于文件元数据的哈希
                assert isinstance(file_hash, str)
                assert len(file_hash) > 0

class TestCacheManagerGetLogic:
    """测试缓存获取逻辑的边界情况。"""

    def test_get_with_memory_cache_hit(self):
        """
        TDD测试：get方法应该正确处理内存缓存命中

        这个测试覆盖第156-157行的内存缓存命中日志
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache'), \
             patch('src.cache.cache_manager.logger') as mock_logger:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value

            # 模拟内存缓存命中
            cached_data = {"test": "data"}
            mock_lru.return_value.get.return_value = cached_data

            # 模拟文件存在
            with patch('src.cache.cache_manager.Path.exists', return_value=True), \
                 patch('src.cache.cache_manager.Path.stat') as mock_stat:
                mock_stat.return_value.st_size = 100
                mock_stat.return_value.st_mtime_ns = 12345

                result = manager.get("test.xlsx")

                # 验证返回了缓存数据
                assert result == cached_data

                # 验证记录了debug日志
                mock_logger.debug.assert_called()
                debug_calls = [call[0][0] for call in mock_logger.debug.call_args_list]
                assert any("Cache hit (memory)" in call for call in debug_calls)

    def test_get_with_disk_cache_hit_and_promotion(self):
        """
        TDD测试：get方法应该正确处理磁盘缓存命中并提升到内存

        这个测试覆盖第164-171行的磁盘缓存逻辑
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk, \
             patch('src.cache.cache_manager.logger') as mock_logger:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟内存缓存未命中，磁盘缓存命中
            mock_lru.return_value.get.return_value = None
            cached_data = {"test": "data", "timestamp": time.time()}
            mock_disk.return_value.get.return_value = cached_data

            # 模拟缓存仍然有效
            with patch.object(manager, '_is_cache_valid', return_value=True):
                # 模拟文件存在
                with patch('src.cache.cache_manager.Path.exists', return_value=True), \
                     patch('src.cache.cache_manager.Path.stat') as mock_stat:
                    mock_stat.return_value.st_size = 100
                    mock_stat.return_value.st_mtime_ns = 12345

                    result = manager.get("test.xlsx")

                    # 验证返回了缓存数据
                    assert result == cached_data

                    # 验证数据被提升到内存缓存
                    mock_lru.return_value.set.assert_called_once()

                    # 验证记录了debug日志
                    mock_logger.debug.assert_called()
                    debug_calls = [call[0][0] for call in mock_logger.debug.call_args_list]
                    assert any("Cache hit (disk)" in call for call in debug_calls)

    def test_get_with_expired_disk_cache(self):
        """
        TDD测试：get方法应该正确处理过期的磁盘缓存

        这个测试覆盖第170-171行的缓存过期逻辑
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk, \
             patch('src.cache.cache_manager.logger') as mock_logger:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟内存缓存未命中，磁盘缓存命中但过期
            mock_lru.return_value.get.return_value = None
            cached_data = {"test": "data", "timestamp": time.time() - 7200}  # 2小时前
            mock_disk.return_value.get.return_value = cached_data

            # 模拟缓存已过期
            with patch.object(manager, '_is_cache_valid', return_value=False):
                # 模拟文件存在
                with patch('src.cache.cache_manager.Path.exists', return_value=True), \
                     patch('src.cache.cache_manager.Path.stat') as mock_stat:
                    mock_stat.return_value.st_size = 100
                    mock_stat.return_value.st_mtime_ns = 12345

                    result = manager.get("test.xlsx")

                    # 验证返回None（缓存未命中）
                    assert result is None

                    # 验证没有提升到内存缓存
                    mock_lru.return_value.set.assert_not_called()

                    # 验证记录了过期日志
                    mock_logger.debug.assert_called()
                    debug_calls = [call[0][0] for call in mock_logger.debug.call_args_list]
                    assert any("Cache expired (disk)" in call for call in debug_calls)

# === 边界情况和错误处理测试 ===

def test_cache_manager_init_with_none_cache_dir():
    """
    TDD测试：CacheManager初始化时应该处理cache_dir为None的情况

    这个测试覆盖第63-64行的代码路径
    """
    mock_config = create_mock_config()
    mock_config.cache_dir = None  # 设置为None

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('src.cache.cache_manager.LRURowBlockCache'), \
         patch('src.cache.cache_manager.DiskCache') as mock_disk_cache, \
         patch('os.path.expanduser', return_value='/home/user'):

        manager = CacheManager(config=mock_config)

        # 验证DiskCache被调用
        mock_disk_cache.assert_called_once()
        call_args = mock_disk_cache.call_args

        # 验证cache_dir参数存在且不为None
        assert 'cache_dir' in call_args[1]
        cache_dir_arg = call_args[1]['cache_dir']
        assert cache_dir_arg is not None

        # 验证代码路径被执行（主要目的是覆盖第63-64行）
        assert manager.disk_cache is not None

def test_cache_manager_set_disk_cache_exception():
    """
    TDD测试：CacheManager.set应该处理磁盘缓存异常

    这个测试覆盖第211-212行的异常处理代码
    """
    mock_config = create_mock_config()

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
         patch('src.cache.cache_manager.DiskCache') as mock_disk, \
         patch('src.cache.cache_manager.logger') as mock_logger, \
         patch('src.cache.cache_manager.Path.exists', return_value=True), \
         patch('src.cache.cache_manager.Path.stat') as mock_stat:

        # 设置文件状态
        mock_stat.return_value.st_size = 100
        mock_stat.return_value.st_mtime_ns = 12345

        manager = CacheManager(config=mock_config)
        manager.memory_cache = mock_lru.return_value
        manager.disk_cache = mock_disk.return_value

        # 模拟磁盘缓存抛出异常
        mock_disk.return_value.set.side_effect = Exception("磁盘写入失败")

        # 调用set方法，应该不抛出异常
        manager.set("test.xlsx", "A1:B2", {"data": "test"})

        # 验证内存缓存仍然被调用
        mock_lru.return_value.set.assert_called_once()

        # 验证记录了磁盘缓存失败的警告日志
        mock_logger.warning.assert_called()
        warning_calls = [call[0][0] for call in mock_logger.warning.call_args_list]
        assert any("Failed to cache on disk" in call for call in warning_calls)

def test_is_cache_valid_with_invalid_cache_entry():
    """
    TDD测试：_is_cache_valid应该处理无效的缓存条目

    这个测试覆盖第223-224行的验证逻辑
    """
    mock_config = create_mock_config()

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config):
        manager = CacheManager(config=mock_config)

        # 测试非字典类型
        assert manager._is_cache_valid("not a dict") is False
        assert manager._is_cache_valid(None) is False
        assert manager._is_cache_valid(123) is False

        # 测试缺少timestamp的字典
        assert manager._is_cache_valid({"data": "test"}) is False
        assert manager._is_cache_valid({}) is False

def test_is_cache_valid_with_expired_cache():
    """
    TDD测试：_is_cache_valid应该检测过期的缓存

    这个测试覆盖第230-231行的过期检查代码
    """
    mock_config = create_mock_config()
    mock_config.cache_expiry_seconds = 3600  # 1小时过期

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('time.time', return_value=7200):  # 当前时间为7200秒

        manager = CacheManager(config=mock_config)

        # 创建过期的缓存条目（时间戳为0，已过期超过1小时）
        expired_entry = {
            "timestamp": 0,
            "data": "test"
        }

        assert manager._is_cache_valid(expired_entry) is False

def test_is_cache_valid_with_missing_file():
    """
    TDD测试：_is_cache_valid应该检测文件不存在的情况

    这个测试覆盖第238-239行的文件存在检查代码
    """
    mock_config = create_mock_config()
    mock_config.cache_expiry_seconds = 3600

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('time.time', return_value=1000), \
         patch('src.cache.cache_manager.Path') as mock_path:

        manager = CacheManager(config=mock_config)

        # 模拟文件不存在
        mock_path_instance = mock_path.return_value
        mock_path_instance.exists.return_value = False

        # 创建有效时间戳但文件不存在的缓存条目
        cache_entry = {
            "timestamp": 500,  # 未过期
            "file_path": "/path/to/missing/file.xlsx"
        }

        assert manager._is_cache_valid(cache_entry) is False

def test_is_cache_valid_with_file_access_exception():
    """
    TDD测试：_is_cache_valid应该处理文件访问异常

    这个测试覆盖第244-245行的异常处理代码
    """
    mock_config = create_mock_config()
    mock_config.cache_expiry_seconds = 3600

    with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
         patch('time.time', return_value=1000), \
         patch('src.cache.cache_manager.Path') as mock_path:

        manager = CacheManager(config=mock_config)

        # 模拟文件访问抛出异常
        mock_path.side_effect = Exception("文件访问权限错误")

        # 创建有效时间戳的缓存条目
        cache_entry = {
            "timestamp": 500,  # 未过期
            "file_path": "/path/to/file.xlsx"
        }

        assert manager._is_cache_valid(cache_entry) is False

class TestCacheManagerUncoveredCode:
    """TDD测试：缓存管理器未覆盖代码测试"""

    def test_cache_manager_initialization_cache_dir_none_fallback(self):
        """
        TDD测试：缓存管理器应该处理cache_dir为None的回退情况

        覆盖代码行：64 - cache_dir为None时的回退逻辑
        """
        mock_config = create_mock_config()
        mock_config.cache_dir = ""  # 设置为空字符串触发回退逻辑

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('os.environ.get') as mock_env_get, \
             patch('os.path.expanduser') as mock_expanduser, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk_cache:

            # 模拟环境变量存在，返回有效路径
            def env_get_side_effect(key, default=None):
                if key == 'LOCALAPPDATA':
                    return 'C:/Users/test/AppData/Local'
                elif key == 'XDG_CACHE_HOME':
                    return '/home/test/.cache'
                return default

            mock_env_get.side_effect = env_get_side_effect
            mock_expanduser.return_value = '/home/user'

            # 创建缓存管理器
            manager = CacheManager(config=mock_config)

            # 验证使用了环境变量的缓存目录
            mock_disk_cache.assert_called_once()
            call_args = mock_disk_cache.call_args
            assert 'mcp-sheet-parser' in call_args[1]['cache_dir']

    def test_cache_manager_is_cache_valid_all_conditions_true(self):
        """
        TDD测试：_is_cache_valid应该在所有条件都满足时返回True

        覆盖代码行：247 - 所有验证条件都通过时的返回逻辑
        """
        mock_config = create_mock_config()
        mock_config.cache_expiry_seconds = 3600

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('time.time', return_value=2000), \
             patch('src.cache.cache_manager.Path') as mock_path_class:

            manager = CacheManager(config=mock_config)

            # 模拟文件存在且时间戳匹配
            mock_path = MagicMock()
            mock_path.exists.return_value = True
            mock_path.stat.return_value.st_mtime_ns = 1500 * 1_000_000_000  # 转换为纳秒
            mock_path_class.return_value = mock_path

            # 创建有效的缓存条目
            cache_entry = {
                "timestamp": 1000,  # 未过期 (2000 - 1000 = 1000 < 3600)
                "file_path": "/path/to/file.xlsx",
                "file_mtime_ns": 1500 * 1_000_000_000
            }

            # 验证返回True
            assert manager._is_cache_valid(cache_entry) is True

    def test_cache_manager_get_stats_disk_cache_exception(self):
        """
        TDD测试：get_stats应该处理磁盘缓存统计异常

        覆盖代码行：299-300 - 磁盘缓存统计异常处理逻辑
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟磁盘缓存目录访问抛出异常
            manager.disk_cache.cache_dir.glob.side_effect = Exception("磁盘访问错误")

            # 获取统计信息
            stats = manager.get_stats()

            # 验证异常被正确处理
            assert 'disk_cache' in stats
            assert 'error' in stats['disk_cache']
            assert '磁盘访问错误' in stats['disk_cache']['error']


class TestCacheManagerOptimization:
    """TDD测试：缓存管理器优化功能测试"""

    def test_cache_manager_optimize_cache_full_workflow(self):
        """
        TDD测试：optimize_cache应该执行完整的缓存优化工作流

        覆盖代码行：311-349 - 完整的缓存优化逻辑
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk, \
             patch('pickle.load') as mock_pickle_load, \
             patch('builtins.open', mock_open()) as mock_file:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟磁盘缓存文件
            mock_cache_file1 = MagicMock()
            mock_cache_file1.name = 'cache1.cache'
            mock_cache_file2 = MagicMock()
            mock_cache_file2.name = 'cache2.cache'

            manager.disk_cache.cache_dir.glob.return_value = [mock_cache_file1, mock_cache_file2]

            # 模拟第一个文件过期，第二个文件有效
            mock_pickle_load.side_effect = [
                {'timestamp': 100, 'file_path': '/path1'},  # 过期
                {'timestamp': 1900, 'file_path': '/path2'}  # 有效
            ]

            # 模拟_is_cache_valid的返回值
            with patch.object(manager, '_is_cache_valid', side_effect=[False, True]):
                # 执行缓存优化
                result = manager.optimize_cache()

                # 验证结果
                assert 'memory_cache_cleaned' in result
                assert 'disk_cache_cleaned' in result
                assert 'errors' in result
                assert result['disk_cache_cleaned'] == 1  # 一个过期文件被清理
                mock_cache_file1.unlink.assert_called_once()

    def test_cache_manager_optimize_cache_corrupted_file_handling(self):
        """
        TDD测试：optimize_cache应该处理损坏的缓存文件

        覆盖代码行：335-342 - 损坏文件处理逻辑
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk, \
             patch('pickle.load') as mock_pickle_load, \
             patch('builtins.open', mock_open()) as mock_file:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟损坏的缓存文件
            mock_cache_file = MagicMock()
            mock_cache_file.name = 'corrupted.cache'
            manager.disk_cache.cache_dir.glob.return_value = [mock_cache_file]

            # 模拟pickle加载失败（使用特定的异常类型）
            import pickle
            mock_pickle_load.side_effect = pickle.PickleError("Pickle加载失败")

            # 执行缓存优化
            result = manager.optimize_cache()

            # 验证损坏文件被移除
            assert result['disk_cache_cleaned'] == 1
            assert len(result['errors']) == 1
            assert 'Pickle加载失败' in result['errors'][0]
            mock_cache_file.unlink.assert_called_once()

    def test_cache_manager_optimize_cache_unlink_error_handling(self):
        """
        TDD测试：optimize_cache应该处理文件删除失败的情况

        覆盖代码行：341-342 - 文件删除失败处理逻辑
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk, \
             patch('pickle.load') as mock_pickle_load, \
             patch('builtins.open', mock_open()) as mock_file:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟损坏的缓存文件
            mock_cache_file = MagicMock()
            mock_cache_file.name = 'corrupted.cache'
            manager.disk_cache.cache_dir.glob.return_value = [mock_cache_file]

            # 模拟pickle加载失败和文件删除失败
            import pickle
            mock_pickle_load.side_effect = pickle.PickleError("Pickle加载失败")
            mock_cache_file.unlink.side_effect = OSError("文件删除权限错误")

            # 执行缓存优化
            result = manager.optimize_cache()

            # 验证错误被正确记录
            assert result['disk_cache_cleaned'] == 0  # 没有成功删除
            assert len(result['errors']) == 1
            assert 'Failed to remove corrupted cache file' in result['errors'][0]
            assert '文件删除权限错误' in result['errors'][0]

    def test_cache_manager_optimize_cache_unexpected_error_handling(self):
        """
        TDD测试：optimize_cache应该处理意外错误

        覆盖代码行：343-345 - 意外错误处理逻辑
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk, \
             patch('pickle.load') as mock_pickle_load, \
             patch('builtins.open', mock_open()) as mock_file:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟缓存文件
            mock_cache_file = MagicMock()
            mock_cache_file.name = 'test.cache'
            manager.disk_cache.cache_dir.glob.return_value = [mock_cache_file]

            # 模拟意外错误（不是pickle或OSError）
            mock_pickle_load.side_effect = RuntimeError("意外的运行时错误")

            # 执行缓存优化
            result = manager.optimize_cache()

            # 验证意外错误被记录但文件不被删除
            assert result['disk_cache_cleaned'] == 0
            assert len(result['errors']) == 1
            assert 'Unexpected error processing cache file' in result['errors'][0]
            assert '意外的运行时错误' in result['errors'][0]
            mock_cache_file.unlink.assert_not_called()

    def test_cache_manager_optimize_cache_disk_cache_general_error(self):
        """
        TDD测试：optimize_cache应该处理磁盘缓存访问的一般错误

        覆盖代码行：346-347 - 磁盘缓存一般错误处理逻辑
        """
        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟磁盘缓存目录访问失败
            manager.disk_cache.cache_dir.glob.side_effect = Exception("磁盘缓存访问失败")

            # 执行缓存优化
            result = manager.optimize_cache()

            # 验证一般错误被正确处理
            assert result['disk_cache_cleaned'] == 0
            assert len(result['errors']) == 1
            assert 'Failed to optimize disk cache' in result['errors'][0]
            assert '磁盘缓存访问失败' in result['errors'][0]


class TestCacheManagerConcurrency:
    """TDD测试：缓存管理器并发安全测试"""

    def test_cache_manager_concurrent_access(self):
        """
        TDD测试：缓存管理器应该支持并发访问

        覆盖代码行：多线程环境下的缓存操作安全性
        """
        import threading
        import time

        mock_config = create_mock_config()

        with patch('src.cache.cache_manager.get_cache_config', return_value=mock_config), \
             patch('src.cache.cache_manager.LRURowBlockCache') as mock_lru, \
             patch('src.cache.cache_manager.DiskCache') as mock_disk:

            manager = CacheManager(config=mock_config)
            manager.memory_cache = mock_lru.return_value
            manager.disk_cache = mock_disk.return_value

            # 模拟内存缓存操作
            manager.memory_cache.get.return_value = None
            manager.memory_cache.set.return_value = None
            manager.disk_cache.get.return_value = None
            manager.disk_cache.set.return_value = None

            results = []
            errors = []

            def cache_operation(thread_id):
                try:
                    for i in range(10):
                        key = f"thread_{thread_id}_key_{i}"
                        value = f"thread_{thread_id}_value_{i}"

                        # 设置缓存
                        manager.set(key, value, "/fake/path")

                        # 获取缓存
                        result = manager.get(key)
                        results.append((thread_id, i, result))

                        time.sleep(0.001)  # 短暂延迟增加并发冲突概率
                except Exception as e:
                    errors.append((thread_id, str(e)))

            # 创建多个线程
            threads = []
            for i in range(5):
                thread = threading.Thread(target=cache_operation, args=(i,))
                threads.append(thread)

            # 启动所有线程
            for thread in threads:
                thread.start()

            # 等待所有线程完成
            for thread in threads:
                thread.join()

            # 验证没有发生错误
            assert len(errors) == 0, f"并发访问出现错误: {errors}"
            assert len(results) == 50  # 5个线程 × 10次操作
