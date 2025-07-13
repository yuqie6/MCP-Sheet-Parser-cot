"""
Configuration settings for the MCP Sheet Parser.
"""

class AppConfig:
    # Streaming settings
    STREAMING_THRESHOLD_CELLS = 10000
    STREAMING_FILE_SIZE_MB = 10
    STREAMING_CHUNK_SIZE_ROWS = 1000

    # File size thresholds for parsing strategies
    SMALL_FILE_THRESHOLD_CELLS = 1000
    LARGE_FILE_THRESHOLD_CELLS = 10000

    # Large file threshold for streaming summary
    STREAMING_SUMMARY_THRESHOLD_CELLS = 50000

    # Cache settings
    CACHE_ENABLED = True
    CACHE_MAX_SIZE = 128  # Max number of items in cache
    CACHE_TTL_SECONDS = 600  # Time-to-live for cache entries

config = AppConfig()
