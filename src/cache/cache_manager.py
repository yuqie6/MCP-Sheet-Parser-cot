"""
主缓存管理器，结合 LRU 与磁盘缓存。

本模块提供统一的缓存操作接口，
将内存中的 LRU 缓存与可选的磁盘持久化结合。
"""

import hashlib
import time
import logging
import threading
from typing import Any
from pathlib import Path
import os

# 缓存管理常量
SMALL_FILE_THRESHOLD_BYTES = 1024 * 1024  # 1MB，小文件阈值，用于内容哈希
HASH_PREFIX_LENGTH = 16  # 内容哈希前缀长度
CACHE_KEY_LENGTH = 32    # 缓存键长度

from ..unified_config import get_cache_config
from .lru_cache import LRURowBlockCache
from .disk_cache import DiskCache

logger = logging.getLogger(__name__)


class CacheManager:
    """
    统一缓存管理器，结合内存 LRU 与磁盘持久化。
    
    按文件哈希+范围键为行块提供智能缓存，适用于重复工具调用。
    """
    
    def __init__(self, config: Any = None):
        """
        初始化缓存管理器。

        参数：
            config: 可选缓存配置。为 None 时使用统一全局配置。
        """
        # 优先使用统一配置系统
        self.config = config or get_cache_config()
        self.config.validate()
        
        # Initialize caches based on configuration
        self.memory_cache = None
        self.disk_cache = None
        
        if self.config.memory_cache_enabled:
            self.memory_cache = LRURowBlockCache(max_entries=self.config.max_entries)
            logger.info(f"Initialized LRU cache with {self.config.max_entries} entries")
        
        if self.config.disk_cache_enabled:
            cache_dir = self.config.cache_dir
            if not cache_dir or cache_dir is None:
                if os.name == 'nt':
                    cache_base = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
                else:
                    cache_base = os.environ.get('XDG_CACHE_HOME', os.path.expanduser('~/.cache'))
                cache_dir = str(Path(cache_base) / 'mcp-sheet-parser')
            # 这里确保 cache_dir 永远不是 None
            if cache_dir is None:
                cache_dir = str(Path(os.path.expanduser('~')) / 'mcp-sheet-parser')
            self.disk_cache = DiskCache(
                cache_dir=cache_dir,
                max_cache_size_mb=self.config.max_disk_cache_size_mb
            )
            logger.info(f"Initialized disk cache at {cache_dir}")

    def _generate_cache_key(self, file_path: str, range_string: str | None = None,
                           sheet_name: str | None = None) -> str:
        """
        根据文件哈希和范围参数生成缓存键。
        
        参数：
            file_path: 文件路径
            range_string: 可选范围字符串（如 "A1:D10"）
            sheet_name: 可选表名
        返回：
            缓存键字符串
        """
        # Calculate file hash
        file_hash = self._calculate_file_hash(file_path)
        
        # Build cache key components
        key_parts = [file_hash]
        if range_string:
            key_parts.append(f"range:{range_string}")
        if sheet_name:
            key_parts.append(f"sheet:{sheet_name}")
        
        cache_key = "|".join(key_parts)
        return cache_key
    
    def _calculate_file_hash(self, file_path: str) -> str:
        """
        计算文件的 SHA256 哈希值，用于生成缓存键。

        包含文件大小、修改时间，小文件还包含内容哈希，确保缓存一致性。

        参数：
            file_path: 文件路径

        返回：
            文件的 SHA256 哈希值，或计算失败时的错误字符串
        """
        try:
            path = Path(file_path)
            if not path.exists():
                return f"missing:{file_path}"

            # 包含文件大小和修改时间以加速比较
            stat = path.stat()
            # 使用高精度修改时间以提升缓存失效准确性
            file_signature = f"{path.name}:{stat.st_size}:{stat.st_mtime_ns}"

            # 对于小文件，包含内容哈希以确保最大准确性
            if stat.st_size < SMALL_FILE_THRESHOLD_BYTES:
                try:
                    with open(path, 'rb') as f:
                        content = f.read()
                    content_hash = hashlib.sha256(content).hexdigest()[:HASH_PREFIX_LENGTH]
                    file_signature += f":{content_hash}"
                except (OSError, MemoryError) as e:
                    logger.warning(f"Failed to read file content for hash: {e}")
                    # 仅回退到文件元数据

            return hashlib.sha256(file_signature.encode()).hexdigest()[:CACHE_KEY_LENGTH]
        except Exception as e:
            logger.warning(f"Failed to calculate file hash for {file_path}: {e}")
            # 返回唯一错误标识以避免缓存冲突
            return f"error:{abs(hash(file_path))}:{int(time.time())}"

    def get(self, file_path: str, range_string: str | None = None,
            sheet_name: str | None = None) -> Any | None:
        """
        获取缓存数据。
        
        参数：
            file_path: 文件路径
            range_string: 可选范围字符串
            sheet_name: 可选表名
        返回：
            找到则返回缓存数据，否则返回 None
        """
        if not self.config.is_cache_enabled():
            return None
        
        cache_key = self._generate_cache_key(file_path, range_string, sheet_name)
        
        # 优先尝试内存缓存
        if self.memory_cache:
            cached_data = self.memory_cache.get(cache_key)
            if cached_data is not None:
                logger.debug(f"Cache hit (memory): {cache_key}")
                return cached_data
        
        # 再尝试磁盘缓存
        if self.disk_cache:
            cached_data = self.disk_cache.get(cache_key)
            if cached_data is not None:
                # 检查缓存条目是否仍然有效
                if self._is_cache_valid(cached_data):
                    logger.debug(f"Cache hit (disk): {cache_key}")
                    # 提升到内存缓存
                    if self.memory_cache:
                        self.memory_cache.set(cache_key, cached_data)
                    return cached_data
                else:
                    logger.debug(f"Cache expired (disk): {cache_key}")
        
        logger.debug(f"Cache miss: {cache_key}")
        return None

    def set(self, file_path: str, data: Any, range_string: str | None = None,
            sheet_name: str | None = None) -> None:
        """
        缓存数据。
        
        参数：
            file_path: 文件路径
            data: 要缓存的数据
            range_string: 可选范围字符串
            sheet_name: 可选表名
        """
        if not self.config.is_cache_enabled():
            return
        
        cache_key = self._generate_cache_key(file_path, range_string, sheet_name)
        
        # 添加时间戳用于过期判断
        cache_entry = {
            'data': data,
            'timestamp': time.time(),
            'file_path': file_path,
            'range_string': range_string,
            'sheet_name': sheet_name
        }
        
        # 存入内存缓存
        if self.memory_cache:
            self.memory_cache.set(cache_key, cache_entry)
            logger.debug(f"Cached in memory: {cache_key}")
        
        # 存入磁盘缓存
        if self.disk_cache:
            try:
                self.disk_cache.set(cache_key, cache_entry)
                logger.debug(f"Cached on disk: {cache_key}")
            except Exception as e:
                logger.warning(f"Failed to cache on disk: {e}")

    def _is_cache_valid(self, cache_entry: dict[str, Any]) -> bool:
        """
        检查缓存条目是否有效。
        
        参数：
            cache_entry: 待验证的缓存条目
        返回：
            有效返回 True，否则返回 False
        """
        if not isinstance(cache_entry, dict) or 'timestamp' not in cache_entry:
            return False
        
        current_time = time.time()
        entry_time = cache_entry['timestamp']
        
        # 检查过期
        if current_time - entry_time > self.config.cache_expiry_seconds:
            return False
        
        # 检查文件是否仍存在且未被修改
        file_path = cache_entry.get('file_path')
        if file_path:
            try:
                path = Path(file_path)
                if not path.exists():
                    return False
                
                # 当前依赖缓存键中的文件哈希
                # 可在此处添加更复杂的校验
                return True
            except Exception:
                return False
        
        return True
    
    def clear(self) -> None:
        """清除所有缓存。"""
        if self.memory_cache:
            self.memory_cache.clear()
            logger.info("Cleared memory cache")
        
        if self.disk_cache:
            self.disk_cache.clear()
            logger.info("Cleared disk cache")

    def get_stats(self) -> dict[str, Any]:
        """
        获取缓存统计信息。
        
        返回：
            包含缓存统计信息的字典
        """
        stats = {
            'config': {
                'cache_enabled': self.config.cache_enabled,
                'memory_cache_enabled': self.config.memory_cache_enabled,
                'disk_cache_enabled': self.config.disk_cache_enabled,
                'max_entries': self.config.max_entries,
                'max_disk_cache_size_mb': self.config.max_disk_cache_size_mb,
                'cache_expiry_seconds': self.config.cache_expiry_seconds,
                'cache_dir': self.config.cache_dir
            },
            'memory_cache': None,
            'disk_cache': None
        }
        
        # 内存缓存统计
        if self.memory_cache:
            stats['memory_cache'] = {
                'current_size': len(self.memory_cache.cache),
                'max_size': self.memory_cache.cache.maxsize,
                'hits': getattr(self.memory_cache.cache, 'hits', 0),
                'misses': getattr(self.memory_cache.cache, 'misses', 0)
            }
        
        # 磁盘缓存统计
        if self.disk_cache:
            try:
                cache_files = list(self.disk_cache.cache_dir.glob('*.cache'))
                total_size = sum(f.stat().st_size for f in cache_files)
                stats['disk_cache'] = {
                    'cache_files': len(cache_files),
                    'total_size_mb': total_size / (1024 * 1024),
                    'cache_dir': str(self.disk_cache.cache_dir)
                }
            except Exception as e:
                stats['disk_cache'] = {'error': str(e)}
        
        return stats

    def optimize_cache(self) -> dict[str, Any]:
        """
        优化缓存，清理过期条目。

        返回：
            包含优化结果的字典
        """
        import pickle  # 移到函数顶部，避免重复导入

        results = {
            'memory_cache_cleaned': 0,
            'disk_cache_cleaned': 0,
            'errors': []
        }

        # 清理内存缓存（LRU自动处理）
        if self.memory_cache:
            # LRU缓存无法轻易清理过期项，除非遍历所有条目
            pass

        # 清理磁盘缓存
        if self.disk_cache:
            try:
                cache_files = list(self.disk_cache.cache_dir.glob('*.cache'))
                for cache_file in cache_files:
                    try:
                        with open(cache_file, 'rb') as f:
                            cache_entry = pickle.load(f)
                            if not self._is_cache_valid(cache_entry):
                                cache_file.unlink()
                                results['disk_cache_cleaned'] += 1
                    except (pickle.PickleError, EOFError, OSError) as e:
                        # 移除损坏的缓存文件
                        try:
                            cache_file.unlink()
                            results['disk_cache_cleaned'] += 1
                            results['errors'].append(f"Removed corrupted cache file {cache_file.name}: {e}")
                        except OSError as unlink_error:
                            results['errors'].append(f"Failed to remove corrupted cache file {cache_file.name}: {unlink_error}")
                    except Exception as e:
                        # 记录未预期错误但不移除文件
                        results['errors'].append(f"Unexpected error processing cache file {cache_file.name}: {e}")
            except Exception as e:
                results['errors'].append(f"Failed to optimize disk cache: {e}")
        
        return results


# 全局缓存管理器实例（线程安全）
_global_cache_manager = None
_cache_manager_lock = threading.Lock()

def get_cache_manager() -> CacheManager:
    """获取全局缓存管理器实例（线程安全）。"""
    global _global_cache_manager
    if _global_cache_manager is None:
        with _cache_manager_lock:
            # 双重检查锁定模式
            if _global_cache_manager is None:
                _global_cache_manager = CacheManager()
    return _global_cache_manager

def reset_cache_manager() -> None:
    """重置全局缓存管理器实例（线程安全）。"""
    global _global_cache_manager
    with _cache_manager_lock:
        if _global_cache_manager is not None:
            _global_cache_manager.clear()
        _global_cache_manager = None
