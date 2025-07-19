"""
磁盘缓存模块。

本模块为获取的数据提供磁盘缓存能力，
以减少重复请求时的获取开销。
"""

import pickle
import logging
from typing import Any
from pathlib import Path
import hashlib

logger = logging.getLogger(__name__)


class DiskCache:
    """基于磁盘的缓存管理类。"""

    def __init__(self, cache_dir: str, max_cache_size_mb: int = 1024):
        self.cache_dir = Path(cache_dir)
        self.max_cache_size_mb = max_cache_size_mb

        # 确保缓存目录存在
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def _get_cache_file_path(self, key: str) -> Path:
        """根据给定的 key 计算缓存文件路径。"""
        hashed_key = hashlib.sha256(key.encode()).hexdigest()
        return self.cache_dir / f"{hashed_key}.cache"

    def get(self, key: str) -> Any:
        """根据 key 获取缓存的值。"""
        cache_file = self._get_cache_file_path(key)
        if not cache_file.exists():
            return None
        try:
            with open(cache_file, 'rb') as f:
                return pickle.load(f)
        except (pickle.PickleError, EOFError, OSError) as e:
            logger.warning(f"加载缓存文件 {cache_file.name} 失败: {e}")
            # 移除损坏的缓存文件
            try:
                cache_file.unlink()
            except OSError:
                pass  # 无法删除文件时忽略
            return None

    def set(self, key: str, value: Any) -> None:
        """根据 key 缓存一个值。"""
        cache_file = self._get_cache_file_path(key)
        try:
            with open(cache_file, 'wb') as f:
                pickle.dump(value, f)
            self._cleanup_cache()
        except (pickle.PickleError, OSError) as e:
            logger.warning(f"保存缓存文件 {cache_file.name} 失败: {e}")
            # 尝试移除可能损坏的文件
            try:
                if cache_file.exists():
                    cache_file.unlink()
            except OSError:
                pass  # 无法删除文件时忽略

    def _cleanup_cache(self) -> None:
        """当缓存目录超出最大容量时进行清理。"""
        try:
            total_size = sum(f.stat().st_size for f in self.cache_dir.glob('*.cache') if f.is_file())
            max_size_bytes = self.max_cache_size_mb * 1024 * 1024
            if total_size > max_size_bytes:
                # 超出容量时移除旧缓存文件
                for cache_file in sorted(self.cache_dir.glob('*.cache'), key=lambda f: f.stat().st_mtime):
                    try:
                        total_size -= cache_file.stat().st_size
                        cache_file.unlink()
                        if total_size <= max_size_bytes:
                            break
                    except (OSError, PermissionError):
                        # 忽略单个文件的错误，继续处理其他文件
                        continue
        except (OSError, PermissionError):
            # 忽略整体的stat错误
            pass

    def clear(self) -> None:
        """清除所有缓存文件。"""
        for cache_file in self.cache_dir.glob('*.cache'):
            try:
                cache_file.unlink()
            except (OSError, PermissionError):
                # 忽略删除错误，继续删除其他文件
                pass
