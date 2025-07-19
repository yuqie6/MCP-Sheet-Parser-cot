
import pytest
from unittest.mock import MagicMock, patch, AsyncMock
from src.mcp_server.server import create_server, main
from mcp.server import Server

@patch('src.mcp_server.server.register_tools')
def test_create_server(mock_register_tools):
    """Test server creation and tool registration."""
    server = create_server()
    assert isinstance(server, Server)
    mock_register_tools.assert_called_once_with(server)

@pytest.mark.asyncio
@patch('src.mcp_server.server.create_server')
@patch('src.mcp_server.server.stdio_server')
async def test_main(mock_stdio_server, mock_create_server):
    """Test the main server entry point."""
    mock_server = MagicMock(spec=Server)
    mock_create_server.return_value = mock_server
    
    # Mock the async context manager for stdio_server
    mock_stdio_server.return_value.__aenter__.return_value = (AsyncMock(), AsyncMock())

    await main()

    mock_create_server.assert_called_once()
    mock_stdio_server.assert_called_once()
    mock_server.run.assert_called_once()
