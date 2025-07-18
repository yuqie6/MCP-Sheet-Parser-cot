"""
统一配置管理模块

整合所有配置项，避免配置冲突和重复定义。
提供线程安全的配置访问和更新机制。
"""

import threading
from dataclasses import dataclass, field
from pathlib import Path
import os


@dataclass
class UnifiedConfig:
    """统一的配置类，整合所有配置项"""
    
    # 文件处理配置
    max_file_size_mb: int = 500
    large_file_threshold_mb: int = 50
    small_file_threshold_mb: int = 5
    
    # 单元格数量阈值（统一所有相关配置）
    small_file_threshold_cells: int = 200    # 小文件：完整数据
    medium_file_threshold_cells: int = 1000  # 中文件：采样数据
    large_file_threshold_cells: int = 3000   # 大文件：摘要数据
    max_cells_count: int = 1000000
    
    # 流式处理配置
    streaming_threshold_cells: int = 2000
    streaming_file_size_mb: int = 5
    streaming_chunk_size_rows: int = 500
    streaming_summary_threshold_cells: int = 10000
    
    # 采样和摘要配置
    sample_rows_count: int = 20
    sample_include_styles: bool = False
    summary_preview_rows: int = 5
    
    # 缓存配置（统一所有缓存相关设置）
    cache_enabled: bool = True
    cache_max_entries: int = 128
    cache_ttl_seconds: int = 600
    cache_dir: str | None = None

    # 磁盘缓存配置
    disk_cache_enabled: bool = True
    max_disk_cache_size_mb: int = 1024
    disk_cache_format: str = 'pickle'
    
    # 内存缓存配置
    memory_cache_enabled: bool = True
    
    # 性能和超时配置
    max_memory_usage_mb: int = 1024
    default_timeout_seconds: int = 300
    large_file_timeout_seconds: int = 600
    
    # 分页配置
    max_page_size: int = 10000
    default_page_size: int = 100
    
    def __post_init__(self):
        """初始化后处理"""
        if self.cache_dir is None:
            self.cache_dir = self._get_default_cache_dir()
    
    def _get_default_cache_dir(self) -> str:
        """获取默认缓存目录"""
        if os.name == 'nt':  # Windows
            cache_base = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        else:  # Unix-like systems
            cache_base = os.environ.get('XDG_CACHE_HOME', os.path.expanduser('~/.cache'))
        
        cache_dir = Path(cache_base) / 'mcp-sheet-parser'
        cache_dir.mkdir(parents=True, exist_ok=True)
        return str(cache_dir)
    
    def validate(self) -> None:
        """验证配置参数的有效性"""
        if self.max_file_size_mb <= 0:
            raise ValueError("max_file_size_mb must be positive")
        
        if self.cache_max_entries <= 0:
            raise ValueError("cache_max_entries must be positive")
        
        if self.cache_ttl_seconds <= 0:
            raise ValueError("cache_ttl_seconds must be positive")
        
        if self.max_disk_cache_size_mb <= 0:
            raise ValueError("max_disk_cache_size_mb must be positive")
        
        if self.disk_cache_format not in ['pickle', 'parquet']:
            raise ValueError("disk_cache_format must be 'pickle' or 'parquet'")
        
        if not (self.small_file_threshold_cells < self.medium_file_threshold_cells < self.large_file_threshold_cells):
            raise ValueError("File size thresholds must be in ascending order")
    
    def get_cache_dir(self) -> Path:
        """获取缓存目录的Path对象，确保 self.cache_dir 不为 None"""
        cache_dir = self.cache_dir
        if not cache_dir:
            cache_dir = self._get_default_cache_dir()
            self.cache_dir = cache_dir  
        return Path(cache_dir)
    
    def is_cache_enabled(self) -> bool:
        """检查是否启用了任何形式的缓存"""
        return self.cache_enabled and (self.memory_cache_enabled or self.disk_cache_enabled)


class ConfigManager:
    """线程安全的配置管理器"""
    
    def __init__(self):
        self._config = UnifiedConfig()
        self._lock = threading.RLock()  
        self._config.validate()
    
    def get_config(self) -> UnifiedConfig:
        """获取当前配置（线程安全）"""
        with self._lock:
            return self._config
    
    def update_config(self, **kwargs) -> None:
        """更新配置参数（线程安全）"""
        with self._lock:
            # 创建新的配置对象
            config_dict = {}
            for field_info in self._config.__dataclass_fields__.values():
                field_name = field_info.name
                config_dict[field_name] = kwargs.get(field_name, getattr(self._config, field_name))
            
            new_config = UnifiedConfig(**config_dict)
            new_config.validate()
            self._config = new_config
    
    def reset_to_defaults(self) -> None:
        """重置为默认配置（线程安全）"""
        with self._lock:
            self._config = UnifiedConfig()
            self._config.validate()
    
    def get_cache_config(self):
        """获取缓存相关配置，兼容旧的CacheConfig接口"""
        config = self.get_config()
        
        # 创建一个兼容对象
        class CacheConfigCompat:
            def __init__(self, unified_config):
                self.cache_enabled = unified_config.cache_enabled
                self.cache_dir = unified_config.cache_dir
                self.max_entries = unified_config.cache_max_entries
                self.max_disk_cache_size_mb = unified_config.max_disk_cache_size_mb
                self.cache_expiry_seconds = unified_config.cache_ttl_seconds
                self.disk_cache_enabled = unified_config.disk_cache_enabled
                self.memory_cache_enabled = unified_config.memory_cache_enabled
                self.disk_cache_format = unified_config.disk_cache_format
            
            def is_cache_enabled(self):
                return self.cache_enabled and (self.memory_cache_enabled or self.disk_cache_enabled)
            
            def validate(self):
                pass  # 已经在UnifiedConfig中验证过了
        
        return CacheConfigCompat(config)


# 全局配置管理器实例
_global_config_manager = None
_config_manager_lock = threading.Lock()


def get_config_manager() -> ConfigManager:
    """获取全局配置管理器实例（线程安全）"""
    global _global_config_manager
    if _global_config_manager is None:
        with _config_manager_lock:
            if _global_config_manager is None:
                _global_config_manager = ConfigManager()
    return _global_config_manager


def get_config() -> UnifiedConfig:
    """获取当前配置的便捷函数"""
    return get_config_manager().get_config()


def update_config(**kwargs) -> None:
    """更新配置的便捷函数"""
    get_config_manager().update_config(**kwargs)


# 为了向后兼容，提供旧的接口
def get_cache_config():
    """获取缓存配置（向后兼容）"""
    return get_config_manager().get_cache_config()


# 创建一个兼容旧config模块的对象
class LegacyConfigCompat:
    """为了向后兼容而创建的配置对象"""
    
    @property
    def STREAMING_THRESHOLD_CELLS(self):
        return get_config().streaming_threshold_cells
    
    @property
    def STREAMING_FILE_SIZE_MB(self):
        return get_config().streaming_file_size_mb
    
    @property
    def STREAMING_CHUNK_SIZE_ROWS(self):
        return get_config().streaming_chunk_size_rows
    
    @property
    def SMALL_FILE_THRESHOLD_CELLS(self):
        return get_config().small_file_threshold_cells
    
    @property
    def MEDIUM_FILE_THRESHOLD_CELLS(self):
        return get_config().medium_file_threshold_cells
    
    @property
    def LARGE_FILE_THRESHOLD_CELLS(self):
        return get_config().large_file_threshold_cells
    
    @property
    def SAMPLE_ROWS_COUNT(self):
        return get_config().sample_rows_count
    
    @property
    def SAMPLE_INCLUDE_STYLES(self):
        return get_config().sample_include_styles
    
    @property
    def SUMMARY_PREVIEW_ROWS(self):
        return get_config().summary_preview_rows
    
    @property
    def STREAMING_SUMMARY_THRESHOLD_CELLS(self):
        return get_config().streaming_summary_threshold_cells
    
    @property
    def CACHE_ENABLED(self):
        return get_config().cache_enabled
    
    @property
    def CACHE_MAX_SIZE(self):
        return get_config().cache_max_entries
    
    @property
    def CACHE_TTL_SECONDS(self):
        return get_config().cache_ttl_seconds


# 创建兼容对象
config = LegacyConfigCompat()
