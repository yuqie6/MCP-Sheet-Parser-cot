"""
用于行块缓存的 LRU 缓存。
"""

from cachetools import LRUCache

class LRURowBlockCache:
    def __init__(self, max_entries=100):
        self.cache = LRUCache(maxsize=max_entries)

    def get(self, key):
        return self.cache.get(key)

    def set(self, key, value):
        self.cache[key] = value

    def clear(self):
        self.cache.clear()

