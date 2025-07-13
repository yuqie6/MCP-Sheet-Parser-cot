"""
Disk cache module for caching.

This module provides on-disk caching capabilities for retrieved data
to reduce fetching overhead for repeated requests.
"""

import os
import pickle
from typing import Any
from pathlib import Path
import hashlib


class DiskCache:
    """Disk-based cache manager class."""

    def __init__(self, cache_dir: str, max_cache_size_mb: int = 1024):
        self.cache_dir = Path(cache_dir)
        self.max_cache_size_mb = max_cache_size_mb

        # Ensure the cache directory exists
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def _get_cache_file_path(self, key: str) -> Path:
        """Calculate the cache file path for a given key."""
        hashed_key = hashlib.sha256(key.encode()).hexdigest()
        return self.cache_dir / f"{hashed_key}.cache"

    def get(self, key: str) -> Any:
        """Retrieve a cached value by its key."""
        cache_file = self._get_cache_file_path(key)
        if not cache_file.exists():
            return None
        with open(cache_file, 'rb') as f:
            return pickle.load(f)

    def set(self, key: str, value: Any) -> None:
        """Cache a value by its key."""
        cache_file = self._get_cache_file_path(key)
        with open(cache_file, 'wb') as f:
            pickle.dump(value, f)
        self._cleanup_cache()

    def _cleanup_cache(self) -> None:
        """Clean up the cache directory if it's exceeded the max size."""
        total_size = sum(f.stat().st_size for f in self.cache_dir.glob('*.cache') if f.is_file())
        max_size_bytes = self.max_cache_size_mb * 1024 * 1024
        if total_size > max_size_bytes:
            # Remove old cache files if exceeds size
            for cache_file in sorted(self.cache_dir.glob('*.cache'), key=lambda f: f.stat().st_mtime):
                total_size -= cache_file.stat().st_size
                cache_file.unlink()
                if total_size <= max_size_bytes:
                    break

    def clear(self) -> None:
        """Clear all cache files."""
        for cache_file in self.cache_dir.glob('*.cache'):
            cache_file.unlink()
