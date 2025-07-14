"""
MCP Sheet Parser 的配置设置。
"""


class AppConfig:
    # 流式读取设置
    STREAMING_THRESHOLD_CELLS = 2000  # 流式读取的单元格阈值
    STREAMING_FILE_SIZE_MB = 5  # 流式读取的文件大小阈值（MB）
    STREAMING_CHUNK_SIZE_ROWS = 500  # 每块流式读取的行数

    # 解析策略的文件大小阈值（针对LLM上下文优化）
    SMALL_FILE_THRESHOLD_CELLS = 200   # 小文件：完整数据
    MEDIUM_FILE_THRESHOLD_CELLS = 1000  # 中文件：采样数据
    LARGE_FILE_THRESHOLD_CELLS = 3000   # 大文件：摘要数据

    # 采样设置
    SAMPLE_ROWS_COUNT = 20  # 采样行数
    SAMPLE_INCLUDE_STYLES = False  # 采样时是否包含样式

    # 摘要设置
    SUMMARY_PREVIEW_ROWS = 5  # 摘要中的预览行数

    # 流式摘要的大文件阈值
    STREAMING_SUMMARY_THRESHOLD_CELLS = 10000

    # 缓存设置
    CACHE_ENABLED = True  # 是否启用缓存
    CACHE_MAX_SIZE = 128  # 缓存中最大条目数
    CACHE_TTL_SECONDS = 600  # 缓存条目的生存时间（秒）

config = AppConfig()
