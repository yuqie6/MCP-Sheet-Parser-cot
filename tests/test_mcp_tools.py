"""
Test cases for MCP tools functionality.
"""

import pytest
import json
import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

from src.mcp_server.tools import (
    _handle_parse_sheet_to_json,
    _handle_convert_json_to_html,
    _handle_convert_file_to_html,
    _handle_convert_file_to_html_file,
    _handle_get_table_summary,
    _handle_get_sheet_metadata,
    _json_to_sheet
)
from src.services.sheet_service import SheetService
from src.models.table_model import Sheet, Row, Cell, Style


@pytest.fixture
def mock_service():
    """Create a mock SheetService for testing."""
    service = MagicMock(spec=SheetService)
    return service


@pytest.fixture
def sample_sheet():
    """Create a sample sheet for testing."""
    return Sheet(
        name="TestSheet",
        rows=[
            Row(cells=[
                Cell(value="Header1", style=Style(bold=True, font_color="#FF0000")),
                Cell(value="Header2", style=Style(bold=True, font_color="#FF0000"))
            ]),
            Row(cells=[
                Cell(value="Data1"),
                Cell(value="Data2")
            ])
        ]
    )


@pytest.fixture
def sample_json_data():
    """Create sample JSON data for testing."""
    return {
        "metadata": {
            "name": "TestSheet",
            "rows": 2,
            "cols": 2,
            "has_merged_cells": False,
            "merged_cells_count": 0
        },
        "data": [
            {
                "row": 0,
                "cells": [
                    {"col": 0, "value": "Header1", "style_id": "style_1"},
                    {"col": 1, "value": "Header2", "style_id": "style_1"}
                ]
            },
            {
                "row": 1,
                "cells": [
                    {"col": 0, "value": "Data1", "style_id": None},
                    {"col": 1, "value": "Data2", "style_id": None}
                ]
            }
        ],
        "merged_cells": [],
        "styles": {
            "style_1": {
                "bold": True,
                "font_color": "#FF0000"
            }
        }
    }


@pytest.mark.asyncio
async def test_parse_sheet_to_json(mock_service):
    """Test parse_sheet_to_json tool."""
    # Test missing file_path
    with pytest.raises(ValueError, match="file_path is required"):
        await _handle_parse_sheet_to_json({}, mock_service)
    
    # Test non-existent file
    with pytest.raises(FileNotFoundError):
        await _handle_parse_sheet_to_json({"file_path": "nonexistent.xlsx"}, mock_service)
    
    # Test successful parsing
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        result = await _handle_parse_sheet_to_json({"file_path": test_file}, mock_service)
        assert len(result) == 1
        assert "JSON conversion completed successfully" in result[0].text
        assert "JSON Data:" in result[0].text


@pytest.mark.asyncio
async def test_convert_json_to_html(mock_service, sample_json_data):
    """Test convert_json_to_html tool."""
    # Test missing parameters
    with pytest.raises(ValueError, match="json_data is required"):
        await _handle_convert_json_to_html({}, mock_service)
    
    with pytest.raises(ValueError, match="output_path is required"):
        await _handle_convert_json_to_html({"json_data": sample_json_data}, mock_service)
    
    # Test successful conversion
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
        output_path = tmp.name
    
    try:
        result = await _handle_convert_json_to_html({
            "json_data": sample_json_data,
            "output_path": output_path
        }, mock_service)
        
        assert len(result) == 1
        assert "HTML file generated successfully" in result[0].text
        assert output_path in result[0].text
        assert Path(output_path).exists()
        
        # Verify HTML content
        html_content = Path(output_path).read_text(encoding='utf-8')
        assert "TestSheet" in html_content
        assert "Header1" in html_content
        assert "Data1" in html_content
    finally:
        Path(output_path).unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_convert_file_to_html(mock_service):
    """Test convert_file_to_html tool."""
    # Test missing file_path
    with pytest.raises(ValueError, match="file_path is required"):
        await _handle_convert_file_to_html({}, mock_service)
    
    # Test non-existent file
    with pytest.raises(FileNotFoundError):
        await _handle_convert_file_to_html({"file_path": "nonexistent.xlsx"}, mock_service)
    
    # Test successful conversion
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        result = await _handle_convert_file_to_html({"file_path": test_file}, mock_service)
        assert len(result) == 1
        assert "HTML conversion completed" in result[0].text


@pytest.mark.asyncio
async def test_convert_file_to_html_file(mock_service):
    """Test convert_file_to_html_file tool."""
    # Test missing parameters
    with pytest.raises(ValueError, match="file_path is required"):
        await _handle_convert_file_to_html_file({}, mock_service)
    
    with pytest.raises(ValueError, match="output_path is required"):
        await _handle_convert_file_to_html_file({"file_path": "test.xlsx"}, mock_service)
    
    # Test successful conversion
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
            output_path = tmp.name
        
        try:
            result = await _handle_convert_file_to_html_file({
                "file_path": test_file,
                "output_path": output_path
            }, mock_service)
            
            assert len(result) == 1
            assert "HTML file conversion completed successfully" in result[0].text
            assert Path(output_path).exists()
        finally:
            Path(output_path).unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_get_table_summary(mock_service):
    """Test get_table_summary tool."""
    # Test missing file_path
    with pytest.raises(ValueError, match="file_path is required"):
        await _handle_get_table_summary({}, mock_service)
    
    # Test non-existent file
    with pytest.raises(FileNotFoundError):
        await _handle_get_table_summary({"file_path": "nonexistent.xlsx"}, mock_service)
    
    # Test successful summary
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        result = await _handle_get_table_summary({"file_path": test_file}, mock_service)
        assert len(result) == 1
        assert "Table Summary for:" in result[0].text
        assert "Basic Statistics:" in result[0].text
        assert "Sample Data" in result[0].text
        assert "Processing Recommendation:" in result[0].text


@pytest.mark.asyncio
async def test_get_sheet_metadata(mock_service):
    """Test get_sheet_metadata tool."""
    # Test missing file_path
    with pytest.raises(ValueError, match="file_path is required"):
        await _handle_get_sheet_metadata({}, mock_service)
    
    # Test non-existent file
    with pytest.raises(FileNotFoundError):
        await _handle_get_sheet_metadata({"file_path": "nonexistent.xlsx"}, mock_service)
    
    # Test successful metadata
    test_file = "tests/data/sample.xlsx"
    if Path(test_file).exists():
        result = await _handle_get_sheet_metadata({"file_path": test_file}, mock_service)
        assert len(result) == 1
        assert "Sheet Metadata for:" in result[0].text
        assert "File Information:" in result[0].text
        assert "Structure:" in result[0].text
        assert "Styling:" in result[0].text
        assert "Output Estimates:" in result[0].text


def test_json_to_sheet_conversion(sample_json_data):
    """Test JSON to Sheet conversion."""
    sheet = _json_to_sheet(sample_json_data)
    
    # Test basic properties
    assert sheet.name == "TestSheet"
    assert len(sheet.rows) == 2
    assert len(sheet.rows[0].cells) == 2
    
    # Test cell values
    assert sheet.rows[0].cells[0].value == "Header1"
    assert sheet.rows[0].cells[1].value == "Header2"
    assert sheet.rows[1].cells[0].value == "Data1"
    assert sheet.rows[1].cells[1].value == "Data2"
    
    # Test styles
    assert sheet.rows[0].cells[0].style is not None
    assert sheet.rows[0].cells[0].style.bold == True
    assert sheet.rows[0].cells[0].style.font_color == "#FF0000"
    assert sheet.rows[1].cells[0].style is None  # No style for data cells


def test_json_to_sheet_empty_data():
    """Test JSON to Sheet conversion with empty data."""
    empty_json = {
        "metadata": {"name": "Empty", "rows": 0, "cols": 0},
        "data": [],
        "merged_cells": [],
        "styles": {}
    }
    
    sheet = _json_to_sheet(empty_json)
    assert sheet.name == "Empty"
    assert len(sheet.rows) == 0
    assert len(sheet.merged_cells) == 0
