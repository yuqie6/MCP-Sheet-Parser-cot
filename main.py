#!/usr/bin/env python3
"""
MCP 表格解析服务器 - 主入口点

一个模型上下文协议服务器，提供表格解析和HTML转换工具。
支持多种文件格式，提供便捷和专业的工具接口。
"""

import asyncio
from src.mcp_server.server import main as server_main


def main():
    """MCP 表格解析服务器的主入口点。"""
    asyncio.run(server_main())


if __name__ == "__main__":
    main()
