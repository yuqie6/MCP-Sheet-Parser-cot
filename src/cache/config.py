"""
Cache configuration module.

Provides a centralized configuration system for caching behavior.
"""

import os
from pathlib import Path
from typing import Optional
from dataclasses import dataclass


@dataclass
class CacheConfig:
    """Configuration for caching behavior."""
    
    # Enable/disable caching
    cache_enabled: bool = True
    
    # Cache directory
    cache_dir: Optional[str] = None
    
    # Maximum entries in LRU cache
    max_entries: int = 100
    
    # Maximum size for disk cache in MB
    max_disk_cache_size_mb: int = 1024  # 1GB
    
    # Cache expiration time in seconds
    cache_expiry_seconds: int = 24 * 60 * 60  # 24 hours
    
    # Enable/disable disk cache
    disk_cache_enabled: bool = True
    
    # Enable/disable memory cache
    memory_cache_enabled: bool = True
    
    # Cache format for disk cache ('pickle' or 'parquet')
    disk_cache_format: str = 'pickle'
    
    # Custom cache directory
    def __post_init__(self):
        if self.cache_dir is None:
            # Use default cache directory
            self.cache_dir = self._get_default_cache_dir()
    
    def _get_default_cache_dir(self) -> str:
        """Get the default cache directory."""
        # Try to use system cache directory
        if os.name == 'nt':  # Windows
            cache_base = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        else:  # Unix-like systems
            cache_base = os.environ.get('XDG_CACHE_HOME', os.path.expanduser('~/.cache'))
        
        cache_dir = Path(cache_base) / 'mcp-sheet-parser'
        cache_dir.mkdir(parents=True, exist_ok=True)
        return str(cache_dir)
    
    def get_cache_dir(self) -> Path:
        """Get the cache directory as a Path object."""
        return Path(self.cache_dir)
    
    def is_cache_enabled(self) -> bool:
        """Check if any caching is enabled."""
        return self.cache_enabled and (self.memory_cache_enabled or self.disk_cache_enabled)
    
    def validate(self) -> None:
        """Validate configuration parameters."""
        if self.max_entries <= 0:
            raise ValueError("max_entries must be positive")
        
        if self.max_disk_cache_size_mb <= 0:
            raise ValueError("max_disk_cache_size_mb must be positive")
        
        if self.cache_expiry_seconds <= 0:
            raise ValueError("cache_expiry_seconds must be positive")
        
        if self.disk_cache_format not in ['pickle', 'parquet']:
            raise ValueError("disk_cache_format must be 'pickle' or 'parquet'")
        
        # Ensure cache directory exists
        cache_path = self.get_cache_dir()
        cache_path.mkdir(parents=True, exist_ok=True)


# Global cache configuration instance
_global_cache_config = CacheConfig()


def get_cache_config() -> CacheConfig:
    """Get the global cache configuration."""
    return _global_cache_config


def set_cache_config(config: CacheConfig) -> None:
    """Set the global cache configuration."""
    global _global_cache_config
    config.validate()
    _global_cache_config = config


def update_cache_config(**kwargs) -> None:
    """Update specific cache configuration parameters."""
    global _global_cache_config
    
    # Create a new config with updated values
    config_dict = {
        'cache_enabled': kwargs.get('cache_enabled', _global_cache_config.cache_enabled),
        'cache_dir': kwargs.get('cache_dir', _global_cache_config.cache_dir),
        'max_entries': kwargs.get('max_entries', _global_cache_config.max_entries),
        'max_disk_cache_size_mb': kwargs.get('max_disk_cache_size_mb', _global_cache_config.max_disk_cache_size_mb),
        'cache_expiry_seconds': kwargs.get('cache_expiry_seconds', _global_cache_config.cache_expiry_seconds),
        'disk_cache_enabled': kwargs.get('disk_cache_enabled', _global_cache_config.disk_cache_enabled),
        'memory_cache_enabled': kwargs.get('memory_cache_enabled', _global_cache_config.memory_cache_enabled),
        'disk_cache_format': kwargs.get('disk_cache_format', _global_cache_config.disk_cache_format),
    }
    
    new_config = CacheConfig(**config_dict)
    set_cache_config(new_config)
