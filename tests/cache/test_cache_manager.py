
import pytest
from unittest.mock import MagicMock, patch
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
