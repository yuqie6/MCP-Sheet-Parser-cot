
import pytest
import pickle
from unittest.mock import MagicMock, patch, mock_open
from pathlib import Path
import os
from src.cache.disk_cache import DiskCache

@pytest.fixture
def temp_cache_dir(tmp_path):
    """提供一个临时缓存目录的fixture。"""
    return tmp_path

@pytest.fixture
def disk_cache(temp_cache_dir):
    """提供DiskCache实例的fixture。"""
    return DiskCache(cache_dir=str(temp_cache_dir))

def test_disk_cache_set_get(disk_cache):
    """测试设置和获取缓存项。"""
    disk_cache.set("my_key", {"data": "my_value"})
    retrieved = disk_cache.get("my_key")
    assert retrieved == {"data": "my_value"}

def test_disk_cache_get_non_existent(disk_cache):
    """测试获取一个不存在的项。"""
    assert disk_cache.get("non_existent_key") is None

@patch("pickle.load", side_effect=pickle.PickleError)
def test_disk_cache_get_corrupted(mock_pickle_load, disk_cache):
    """测试获取一个已损坏的缓存项。"""
    cache_file = disk_cache._get_cache_file_path("corrupted_key")
    cache_file.touch()
    assert disk_cache.get("corrupted_key") is None

def test_disk_cache_clear(disk_cache):
    """测试清空缓存。"""
    disk_cache.set("key1", "value1")
    disk_cache.set("key2", "value2")
    disk_cache.clear()
    assert disk_cache.get("key1") is None
    assert disk_cache.get("key2") is None

@patch('src.cache.disk_cache.Path.glob')
def test_disk_cache_cleanup(mock_glob, temp_cache_dir):
    """测试缓存清理机制。"""
    # 为测试设置最大缓存为1MB
    cache = DiskCache(cache_dir=str(temp_cache_dir), max_cache_size_mb=1)

    # 使用 os.stat_result 构造更真实的 stat 对象
    # stat_result: (mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime)
    stat1 = os.stat_result((0, 0, 0, 0, 0, 0, 700 * 1024, 0, 1, 0))
    stat2 = os.stat_result((0, 0, 0, 0, 0, 0, 700 * 1024, 0, 2, 0))

    file1 = MagicMock(spec=Path)
    file1.is_file.return_value = True
    file1.stat.return_value = stat1

    file2 = MagicMock(spec=Path)
    file2.is_file.return_value = True
    file2.stat.return_value = stat2
    
    mock_glob.return_value = [file1, file2]

    cache._cleanup_cache()
    
    # 预期最旧的文件 (file1) 会被删除
    file1.unlink.assert_called_once()
    file2.unlink.assert_not_called()


def test_disk_cache_initialization_creates_directory(tmp_path):
    """
    TDD测试：DiskCache初始化应该创建缓存目录

    这个测试覆盖第30-31行的目录创建代码路径
    """
    cache_dir = tmp_path / "new_cache_dir"
    assert not cache_dir.exists()

    disk_cache = DiskCache(cache_dir=str(cache_dir))

    # 验证目录是否被创建
    assert cache_dir.exists()
    assert cache_dir.is_dir()

def test_disk_cache_initialization_with_existing_directory(temp_cache_dir):
    """
    TDD测试：DiskCache初始化应该处理已存在的目录

    这个测试确保已存在的目录不会导致错误
    """
    assert temp_cache_dir.exists()

    # 验证初始化不会抛出异常
    disk_cache = DiskCache(cache_dir=str(temp_cache_dir))
    assert disk_cache.cache_dir == Path(str(temp_cache_dir))

def test_get_cache_file_path(disk_cache):
    """
    TDD测试：_get_cache_file_path应该生成正确的文件路径

    这个测试覆盖第34-35行的文件路径生成代码路径
    """
    key = "test_key"
    file_path = disk_cache._get_cache_file_path(key)

    assert file_path.parent == disk_cache.cache_dir
    assert file_path.suffix == ".cache"
    # 验证文件名是key的SHA256哈希值
    import hashlib
    expected_hash = hashlib.sha256(key.encode()).hexdigest()
    assert file_path.name == f"{expected_hash}.cache"

def test_set_with_pickle_error(disk_cache):
    """
    TDD测试：set应该处理pickle序列化错误

    这个测试覆盖第44-45行的pickle错误处理代码路径
    """
    # 创建一个不能被pickle序列化的对象
    class UnpicklableObject:
        def __reduce__(self):
            raise pickle.PickleError("此对象无法被pickle序列化")

    unpicklable = UnpicklableObject()

    # set操作应静默失败，不抛出异常
    disk_cache.set("unpicklable_key", unpicklable)

    # 获取时应返回None
    assert disk_cache.get("unpicklable_key") is None

@patch("builtins.open", side_effect=IOError("File write error"))
def test_set_with_io_error(mock_open, disk_cache):
    """
    TDD测试：set应该处理文件写入错误

    这个测试覆盖第46-47行的IO错误处理代码路径
    """
    # set操作应静默失败，不抛出异常
    disk_cache.set("io_error_key", {"data": "value"})

    # 由于写入失败，获取时应返回None
    # 注意：此处需重新创建disk_cache实例以避免mock泄露影响get操作
    new_cache = DiskCache(cache_dir=str(disk_cache.cache_dir))
    assert new_cache.get("io_error_key") is None

def test_get_with_file_not_found(disk_cache):
    """
    TDD测试：get应该处理文件不存在的情况

    这个测试覆盖第52行的文件不存在处理代码路径
    """
    # 获取一个不存在的键
    result = disk_cache.get("non_existent_key")

    # 结果应为None
    assert result is None

def test_get_with_io_error(disk_cache):
    """
    TDD测试：get应该处理文件读取错误

    这个测试覆盖第58-59行的IO错误处理代码路径
    """
    # 创建一个实际的缓存文件用于测试
    cache_file = disk_cache._get_cache_file_path("io_error_key")
    cache_file.touch()

    # 模拟文件读取时发生IOError
    with patch("builtins.open", side_effect=IOError("文件读取错误")):
        result = disk_cache.get("io_error_key")

    # 结果应为None
    assert result is None

def test_clear_with_no_files(disk_cache):
    """
    TDD测试：clear应该处理空缓存目录

    这个测试确保clear在没有文件时不会出错
    """
    # 确保缓存目录为空
    assert len(list(disk_cache.cache_dir.glob("*.cache"))) == 0

    # clear操作不应抛出异常
    disk_cache.clear()

def test_clear_with_permission_error(disk_cache):
    """
    TDD测试：clear应该处理文件删除权限错误

    这个测试覆盖第68-69行的权限错误处理代码路径
    """
    # 创建一个缓存文件
    disk_cache.set("permission_test", {"data": "value"})

    # 模拟删除文件时发生PermissionError
    with patch.object(Path, 'unlink', side_effect=PermissionError("权限不足")):
        # clear操作应静默失败，不抛出异常
        disk_cache.clear()

def test_cleanup_cache_with_no_files(disk_cache):
    """
    TDD测试：_cleanup_cache应该处理没有缓存文件的情况

    这个测试确保cleanup在没有文件时正确处理
    """
    # 确保缓存目录中没有文件
    assert len(list(disk_cache.cache_dir.glob("*.cache"))) == 0

    # 清理操作不应抛出异常
    disk_cache._cleanup_cache()

def test_cleanup_cache_under_size_limit(disk_cache):
    """
    TDD测试：_cleanup_cache应该在大小限制内时不删除文件

    这个测试覆盖第76行的大小检查代码路径
    """
    # 创建几个小文件
    disk_cache.set("small1", "data1")
    disk_cache.set("small2", "data2")

    files_before = len(list(disk_cache.cache_dir.glob("*.cache")))

    disk_cache._cleanup_cache()

    # 文件数量应保持不变（因为总大小未超出限制）
    files_after = len(list(disk_cache.cache_dir.glob("*.cache")))
    assert files_after == files_before

def test_cleanup_cache_with_stat_error(disk_cache):
    """
    TDD测试：_cleanup_cache应该处理文件stat错误

    这个测试覆盖第82-83行的stat错误处理代码路径
    """
    # 创建一个缓存文件
    disk_cache.set("stat_error_test", {"data": "value"})

    # 模拟获取文件状态时发生OSError
    with patch.object(Path, 'stat', side_effect=OSError("获取状态失败")):
        # 清理操作应处理异常且不抛出
        disk_cache._cleanup_cache()

def test_cleanup_cache_with_unlink_error(disk_cache):
    """
    TDD测试：_cleanup_cache应该处理文件删除错误

    这个测试覆盖第90-91行的unlink错误处理代码路径
    """
    # 创建多个大文件以触发清理机制
    large_data = "x" * 1000
    for i in range(5):
        disk_cache.set(f"large_file_{i}", large_data)

    # 模拟删除文件时发生OSError
    with patch.object(Path, 'unlink', side_effect=OSError("删除失败")):
        # 清理操作应处理异常且不抛出
        disk_cache._cleanup_cache()

def test_disk_cache_key_sanitization(disk_cache):
    """
    TDD测试：DiskCache应该处理特殊字符的键名

    这个测试确保特殊字符在文件名中被正确处理
    """
    # 使用包含特殊字符的键名
    special_key = "key/with\\special:chars*and?quotes"
    test_data = {"test": "data"}

    # 验证可以成功设置和获取
    disk_cache.set(special_key, test_data)
    result = disk_cache.get(special_key)

    assert result == test_data

@patch("pathlib.Path.unlink", side_effect=OSError("Permission denied"))
def test_get_with_corrupted_file_unlink_error(mock_unlink, disk_cache):
    """
    TDD测试：get应该处理删除损坏文件时的OSError

    这个测试覆盖第45-46行的OSError处理代码
    """
    # 创建一个已损坏的缓存文件
    cache_file = disk_cache._get_cache_file_path("corrupted_key")
    cache_file.write_bytes(b"corrupted data")

    # 模拟pickle.load抛出异常以触发删除逻辑
    with patch("pickle.load", side_effect=pickle.PickleError("数据损坏")):
        result = disk_cache.get("corrupted_key")

        # 即使删除文件失败，也应返回None
        assert result is None

        # 验证是否尝试了删除操作
        mock_unlink.assert_called_once()

@patch("pathlib.Path.unlink", side_effect=OSError("Permission denied"))
def test_set_with_existing_file_unlink_error(mock_unlink, disk_cache):
    """
    TDD测试：set应该处理删除现有文件时的OSError

    这个测试覆盖第62-63行的OSError处理代码
    """
    # 首先设置一个初始值
    disk_cache.set("test_key", "original_value")

    # 模拟pickle.dump抛出异常以触发删除逻辑
    with patch("pickle.dump", side_effect=pickle.PickleError("序列化失败")):
        # 尝试用新值覆盖，这将触发删除现有文件的逻辑
        disk_cache.set("test_key", "new_value")

        # 验证是否尝试了删除操作
        mock_unlink.assert_called()

def test_cleanup_cache_with_files_exceeding_limit(temp_cache_dir):
    """
    TDD测试：_cleanup_cache应该删除超出限制的文件

    这个测试覆盖第72-80行的文件删除逻辑
    """
    # 创建一个小容量的缓存实例
    disk_cache = DiskCache(cache_dir=str(temp_cache_dir), max_cache_size_mb=0.001)  # 1KB 大小限制

    # 创建多个文件以超出缓存限制
    large_data = "x" * 500
    disk_cache.set("file1", large_data)
    disk_cache.set("file2", large_data)
    disk_cache.set("file3", large_data)  # 此处应触发清理

    # 验证部分文件已被删除
    cache_files = list(temp_cache_dir.glob("*.cache"))
    # 清理后，文件数量应少于3个
    assert len(cache_files) < 3


@patch("pathlib.Path.glob", side_effect=OSError("Glob failed"))
def test_cleanup_cache_with_overall_error(mock_glob, temp_cache_dir):
    """
    TDD测试：_cleanup_cache应该处理整体的OSError

    这个测试覆盖第81-83行的整体错误处理代码
    """
    disk_cache = DiskCache(cache_dir=str(temp_cache_dir))

    # 触发清理，操作不应抛出异常
    disk_cache._cleanup_cache()

    # 验证glob方法被调用
    mock_glob.assert_called()
