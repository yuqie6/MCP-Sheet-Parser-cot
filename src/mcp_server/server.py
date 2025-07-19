#!/usr/bin/env python3
"""
MCP 表格解析服务器

一个模型上下文协议服务器，提供表格解析和HTML转换工具。
支持多种文件格式，提供便捷和专业的工具接口。
"""

import asyncio
import logging

from mcp.server import Server
from mcp.server.stdio import stdio_server

from ..models.tools import register_tools

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def create_server() -> Server:
    """创建并配置MCP服务器。"""
    server = Server("mcp-sheet-parser")

    # 注册所有工具
    register_tools(server)

    logger.info("MCP 表格解析服务器初始化完成")
    return server


async def main():
    """MCP 表格解析服务器的主入口点。"""
    try:
        server = create_server()

        # 使用stdio传输运行服务器
        async with stdio_server() as (read_stream, write_stream):
            logger.info("启动 MCP 表格解析服务器...")
            await server.run(
                read_stream,
                write_stream,
                server.create_initialization_options()
            )
    except Exception as e:
        logger.error(f"服务器错误: {e}")
        raise


if __name__ == "__main__":
    asyncio.run(main())
