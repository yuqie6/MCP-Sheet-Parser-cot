"""
Main cache manager that combines LRU and disk caching.

This module provides a unified interface for caching operations,
combining in-memory LRU cache with optional disk-based persistence.
"""

import hashlib
import time
import logging
from typing import Any, Dict, Optional, Tuple
from pathlib import Path
import pandas as pd

from .config import get_cache_config, CacheConfig
from .lru_cache import LRURowBlockCache
from .disk_cache import DiskCache

logger = logging.getLogger(__name__)


class CacheManager:
    """
    Unified cache manager that combines LRU in-memory cache with disk persistence.
    
    The cache manager provides intelligent caching for row blocks keyed by
    file hash + range for repeated tool calls.
    """
    
    def __init__(self, config: Optional[CacheConfig] = None):
        """
        Initialize the cache manager.
        
        Args:
            config: Optional cache configuration. If None, uses global config.
        """
        self.config = config or get_cache_config()
        self.config.validate()
        
        # Initialize caches based on configuration
        self.memory_cache = None
        self.disk_cache = None
        
        if self.config.memory_cache_enabled:
            self.memory_cache = LRURowBlockCache(max_entries=self.config.max_entries)
            logger.info(f"Initialized LRU cache with {self.config.max_entries} entries")
        
        if self.config.disk_cache_enabled:
            self.disk_cache = DiskCache(
                cache_dir=self.config.cache_dir,
                max_cache_size_mb=self.config.max_disk_cache_size_mb
            )
            logger.info(f"Initialized disk cache at {self.config.cache_dir}")
    
    def _generate_cache_key(self, file_path: str, range_string: Optional[str] = None,
                           sheet_name: Optional[str] = None) -> str:
        """
        Generate a cache key based on file hash and range parameters.
        
        Args:
            file_path: Path to the file
            range_string: Optional range string (e.g., "A1:D10")
            sheet_name: Optional sheet name
            
        Returns:
            Cache key string
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
        Calculate SHA256 hash of the file for cache key generation.
        
        Args:
            file_path: Path to the file
            
        Returns:
            SHA256 hash of the file
        """
        try:
            path = Path(file_path)
            if not path.exists():
                return f"missing:{file_path}"
            
            # Include file size and modification time for faster comparison
            stat = path.stat()
            file_signature = f"{path.name}:{stat.st_size}:{stat.st_mtime}"
            
            # For small files, include content hash
            if stat.st_size < 1024 * 1024:  # 1MB
                with open(path, 'rb') as f:
                    content = f.read()
                content_hash = hashlib.sha256(content).hexdigest()[:16]
                file_signature += f":{content_hash}"
            
            return hashlib.sha256(file_signature.encode()).hexdigest()[:32]
        except Exception as e:
            logger.warning(f"Failed to calculate file hash for {file_path}: {e}")
            return f"error:{file_path}"
    
    def get(self, file_path: str, range_string: Optional[str] = None,
            sheet_name: Optional[str] = None) -> Optional[Any]:
        """
        Retrieve cached data.
        
        Args:
            file_path: Path to the file
            range_string: Optional range string
            sheet_name: Optional sheet name
            
        Returns:
            Cached data if found, None otherwise
        """
        if not self.config.is_cache_enabled():
            return None
        
        cache_key = self._generate_cache_key(file_path, range_string, sheet_name)
        
        # Try memory cache first
        if self.memory_cache:
            cached_data = self.memory_cache.get(cache_key)
            if cached_data is not None:
                logger.debug(f"Cache hit (memory): {cache_key}")
                return cached_data
        
        # Try disk cache
        if self.disk_cache:
            cached_data = self.disk_cache.get(cache_key)
            if cached_data is not None:
                # Check if cache entry is still valid
                if self._is_cache_valid(cached_data):
                    logger.debug(f"Cache hit (disk): {cache_key}")
                    # Promote to memory cache
                    if self.memory_cache:
                        self.memory_cache.set(cache_key, cached_data)
                    return cached_data
                else:
                    logger.debug(f"Cache expired (disk): {cache_key}")
        
        logger.debug(f"Cache miss: {cache_key}")
        return None
    
    def set(self, file_path: str, data: Any, range_string: Optional[str] = None,
            sheet_name: Optional[str] = None) -> None:
        """
        Cache data.
        
        Args:
            file_path: Path to the file
            data: Data to cache
            range_string: Optional range string
            sheet_name: Optional sheet name
        """
        if not self.config.is_cache_enabled():
            return
        
        cache_key = self._generate_cache_key(file_path, range_string, sheet_name)
        
        # Add timestamp for expiration
        cache_entry = {
            'data': data,
            'timestamp': time.time(),
            'file_path': file_path,
            'range_string': range_string,
            'sheet_name': sheet_name
        }
        
        # Store in memory cache
        if self.memory_cache:
            self.memory_cache.set(cache_key, cache_entry)
            logger.debug(f"Cached in memory: {cache_key}")
        
        # Store in disk cache
        if self.disk_cache:
            try:
                self.disk_cache.set(cache_key, cache_entry)
                logger.debug(f"Cached on disk: {cache_key}")
            except Exception as e:
                logger.warning(f"Failed to cache on disk: {e}")
    
    def _is_cache_valid(self, cache_entry: Dict[str, Any]) -> bool:
        """
        Check if a cache entry is still valid.
        
        Args:
            cache_entry: Cache entry to validate
            
        Returns:
            True if valid, False otherwise
        """
        if not isinstance(cache_entry, dict) or 'timestamp' not in cache_entry:
            return False
        
        current_time = time.time()
        entry_time = cache_entry['timestamp']
        
        # Check expiration
        if current_time - entry_time > self.config.cache_expiry_seconds:
            return False
        
        # Check if file still exists and hasn't been modified
        file_path = cache_entry.get('file_path')
        if file_path:
            try:
                path = Path(file_path)
                if not path.exists():
                    return False
                
                # For now, we rely on file hash in the cache key
                # More sophisticated validation could be added here
                return True
            except Exception:
                return False
        
        return True
    
    def clear(self) -> None:
        """Clear all caches."""
        if self.memory_cache:
            self.memory_cache.clear()
            logger.info("Cleared memory cache")
        
        if self.disk_cache:
            self.disk_cache.clear()
            logger.info("Cleared disk cache")
    
    def get_stats(self) -> Dict[str, Any]:
        """
        Get cache statistics.
        
        Returns:
            Dictionary with cache statistics
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
        
        # Memory cache stats
        if self.memory_cache:
            stats['memory_cache'] = {
                'current_size': len(self.memory_cache.cache),
                'max_size': self.memory_cache.cache.maxsize,
                'hits': getattr(self.memory_cache.cache, 'hits', 0),
                'misses': getattr(self.memory_cache.cache, 'misses', 0)
            }
        
        # Disk cache stats
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
    
    def optimize_cache(self) -> Dict[str, Any]:
        """
        Optimize cache by cleaning up expired entries.
        
        Returns:
            Dictionary with optimization results
        """
        results = {
            'memory_cache_cleaned': 0,
            'disk_cache_cleaned': 0,
            'errors': []
        }
        
        # Clean up memory cache (LRU handles this automatically)
        if self.memory_cache:
            # We can't easily clean expired entries from LRU cache
            # without iterating through all entries
            pass
        
        # Clean up disk cache
        if self.disk_cache:
            try:
                cache_files = list(self.disk_cache.cache_dir.glob('*.cache'))
                for cache_file in cache_files:
                    try:
                        with open(cache_file, 'rb') as f:
                            import pickle
                            cache_entry = pickle.load(f)
                            if not self._is_cache_valid(cache_entry):
                                cache_file.unlink()
                                results['disk_cache_cleaned'] += 1
                    except Exception as e:
                        # Remove corrupted cache files
                        cache_file.unlink()
                        results['disk_cache_cleaned'] += 1
                        results['errors'].append(f"Removed corrupted cache file: {e}")
            except Exception as e:
                results['errors'].append(f"Failed to optimize disk cache: {e}")
        
        return results


# Global cache manager instance
_global_cache_manager = None


def get_cache_manager() -> CacheManager:
    """Get the global cache manager instance."""
    global _global_cache_manager
    if _global_cache_manager is None:
        _global_cache_manager = CacheManager()
    return _global_cache_manager


def reset_cache_manager() -> None:
    """Reset the global cache manager instance."""
    global _global_cache_manager
    _global_cache_manager = None
