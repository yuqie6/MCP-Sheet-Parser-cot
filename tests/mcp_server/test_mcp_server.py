import pytest
from unittest.mock import patch, MagicMock
from mcp.server import Server
from src.mcp_server.server import create_server

@patch('src.mcp_server.server.register_tools')
def test_create_server(mock_register_tools: MagicMock):
    """
    测试 create_server 函数是否能正确创建和配置服务器。
    """
    # 调用函数
    server = create_server()

    # 验证返回的是 Server 实例
    assert isinstance(server, Server)

    # 验证服务器名称
    assert server.name == "mcp-sheet-parser"

    # 验证 register_tools 是否被调用
    mock_register_tools.assert_called_once_with(server)