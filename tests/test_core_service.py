import pytest
from unittest.mock import MagicMock, patch, mock_open
from pathlib import Path
from src.core_service import CoreService
from src.models.table_model import Sheet, Row, Cell
from src.exceptions import FileNotFoundError

@pytest.fixture
def mock_parser_factory():
    """Fixture for a mocked ParserFactory."""
    parser = MagicMock()
    sheet = Sheet(name="TestSheet", rows=[Row(cells=[Cell(value="Header")]), Row(cells=[Cell(value="data")])])
    parser.parse.return_value = [sheet]
    factory = MagicMock()
    factory.get_parser.return_value = parser
    return factory

@pytest.fixture
def core_service(mock_parser_factory):
    """Fixture for CoreService with mocked dependencies."""
    with patch('src.core_service.ParserFactory', return_value=mock_parser_factory):
        service = CoreService()
    return service

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
@patch('src.core_service.get_cache_manager')
def test_parse_sheet_simple(mock_get_cache, mock_validate, core_service):
    """Test simple sheet parsing."""
    mock_cache = MagicMock()
    mock_cache.get.return_value = None
    mock_get_cache.return_value = mock_cache

    result = core_service.parse_sheet("dummy.xlsx")
    assert result['sheet_name'] == 'TestSheet'
    assert len(result['rows']) > 0

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
@patch('src.core_service.get_cache_manager')
def test_parse_sheet_from_cache(mock_get_cache, mock_validate, core_service):
    """Test parsing from cache."""
    mock_cache = MagicMock()
    cached_data = {'data': {'sheet_name': 'CachedSheet'}}
    mock_cache.get.return_value = cached_data
    mock_get_cache.return_value = mock_cache

    result = core_service.parse_sheet("dummy.xlsx")
    assert result['sheet_name'] == 'CachedSheet'

@patch('src.core_service.Path.exists', return_value=True)
@patch('src.core_service.HTMLConverter')
def test_convert_to_html(mock_html_converter, mock_path_exists, core_service):
    """Test HTML conversion."""
    core_service.convert_to_html("dummy.xlsx")
    mock_html_converter.return_value.convert_to_files.assert_called_once()

@patch('src.core_service.Path.exists', return_value=True)
@patch('src.converters.paginated_html_converter.PaginatedHTMLConverter')
def test_convert_to_html_paginated(mock_paginated_converter, mock_path_exists, core_service):
    """Test paginated HTML conversion."""
    core_service.convert_to_html("dummy.xlsx", page_size=10)
    mock_paginated_converter.return_value.convert_to_file.assert_called_once()

@patch('src.core_service.Path.exists', return_value=True)
@patch('src.core_service.CoreService._write_back_xlsx')
@patch('shutil.copy2')
def test_apply_changes_xlsx(mock_copy, mock_write_back, mock_path_exists, core_service):
    """Test applying changes to an XLSX file."""
    changes = {"sheet_name": "Sheet1", "headers": [], "rows": []}
    core_service.apply_changes("dummy.xlsx", changes)
    mock_write_back.assert_called_once()
    mock_copy.assert_called_once()

@patch('src.core_service.Path.exists', return_value=True)
@patch('src.core_service.CoreService._write_back_csv')
@patch('shutil.copy2')
def test_apply_changes_csv(mock_copy, mock_write_back, mock_path_exists, core_service):
    """Test applying changes to a CSV file."""
    changes = {"sheet_name": "Sheet1", "headers": [], "rows": []}
    core_service.apply_changes("dummy.csv", changes)
    mock_write_back.assert_called_once()
    mock_copy.assert_called_once()

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
def test_parse_sheet_optimized(mock_validate, core_service):
    """Test optimized sheet parsing."""
    result = core_service.parse_sheet_optimized("dummy.xlsx", include_full_data=True, include_styles=True)
    assert result['sheet_name'] == 'TestSheet'
    assert len(result['rows']) > 0

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
def test_parse_sheet_with_range(mock_validate, core_service):
    """Test parsing with a range string."""
    with patch('src.core_service.parse_range_string', return_value=(0, 0, 0, 0)):
        result = core_service.parse_sheet_optimized("dummy.xlsx", range_string="A1:A1")
        assert result['sheet_name'] == 'TestSheet'

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
@patch('src.core_service.get_cache_manager')
def test_parse_sheet_with_non_existing_sheet_name(mock_get_cache, mock_validate, core_service):
    """Test parsing with a non-existing sheet name."""
    mock_cache = MagicMock()
    mock_cache.get.return_value = None
    mock_get_cache.return_value = mock_cache

    with pytest.raises(ValueError, match="工作表 'NonExistingSheet' 不存在。"):
        core_service.parse_sheet("dummy.xlsx", sheet_name="NonExistingSheet")

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
@patch('src.core_service.get_cache_manager')
def test_parse_sheet_with_no_sheets(mock_get_cache, mock_validate, core_service):
    """Test parsing a file with no sheets."""
    mock_cache = MagicMock()
    mock_cache.get.return_value = None
    mock_get_cache.return_value = mock_cache
    
    # Configure the mock parser to return an empty list of sheets
    core_service.parser_factory.get_parser.return_value.parse.return_value = []

    with pytest.raises(ValueError, match="文件中没有找到任何工作表。"):
        core_service.parse_sheet("dummy.xlsx")

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
@patch('src.core_service.get_cache_manager')
def test_parse_empty_sheet(mock_get_cache, mock_validate, core_service):
    """Test parsing an empty sheet."""
    mock_cache = MagicMock()
    mock_cache.get.return_value = None
    mock_get_cache.return_value = mock_cache

    # Configure the mock parser to return a sheet with no rows
    empty_sheet = Sheet(name="EmptySheet", rows=[])
    core_service.parser_factory.get_parser.return_value.parse.return_value = [empty_sheet]

    result = core_service.parse_sheet("dummy.xlsx")
    assert result['sheet_name'] == 'EmptySheet'
    assert result['rows'] == []
    assert result['total_rows'] == 0


@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
@patch('src.core_service.get_cache_manager')
@patch('src.core_service.CoreService._should_use_streaming', return_value=True)
@patch('src.core_service.CoreService._parse_sheet_streaming')
def test_parse_sheet_streaming_enabled(mock_parse_streaming, mock_should_stream, mock_get_cache, mock_validate, core_service):
    """Test that streaming is called when it should be."""
    mock_cache = MagicMock()
    mock_cache.get.return_value = None
    mock_get_cache.return_value = mock_cache
    mock_parse_streaming.return_value = {"sheet_name": "StreamedSheet"}

    result = core_service.parse_sheet("dummy.xlsx")

    mock_should_stream.assert_called_once()
    mock_parse_streaming.assert_called_once()
    assert result["sheet_name"] == "StreamedSheet"


@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
def test_parse_sheet_optimized_non_existing_sheet(mock_validate, core_service):
    """Test optimized parsing with a non-existing sheet name."""
    with pytest.raises(ValueError, match="工作表 'NonExistingSheet' 不存在。"):
        core_service.parse_sheet_optimized("dummy.xlsx", sheet_name="NonExistingSheet")

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
def test_parse_sheet_optimized_no_sheets(mock_validate, core_service):
    """Test optimized parsing with no sheets in the file."""
    core_service.parser_factory.get_parser.return_value.parse.return_value = []
    with pytest.raises(ValueError, match="文件中没有找到任何工作表。"):
        core_service.parse_sheet_optimized("dummy.xlsx")

@patch('src.core_service.validate_file_input', return_value=(Path("dummy.xlsx"), "xlsx"))
def test_parse_sheet_optimized_invalid_range(mock_validate, core_service):
    """Test optimized parsing with an invalid range string."""
    with pytest.raises(ValueError, match="范围格式错误"):
        core_service.parse_sheet_optimized("dummy.xlsx", range_string="INVALID_RANGE")


@patch('src.core_service.Path')
def test_convert_to_html_file_not_found(mock_path, core_service):
    """Test HTML conversion when the source file does not exist."""
    mock_path.return_value.exists.return_value = False
    with pytest.raises(FileNotFoundError):
        core_service.convert_to_html("non_existing_file.xlsx")

@patch('src.core_service.HTMLConverter')
@patch('src.core_service.Path')
def test_convert_to_html_default_output_path(mock_path, mock_html_converter, core_service):
    """Test HTML conversion with a default output path."""
    mock_path.return_value.exists.return_value = True
    mock_path.return_value.with_suffix.return_value = "dummy.html"
    core_service.convert_to_html("dummy.xlsx")
    # Verify that the output path is correctly defaulted
    # The first argument of the call is the list of sheets, the second is the output path
    call_args = mock_html_converter.return_value.convert_to_files.call_args
    output_path = call_args[0][1]
    assert str(output_path) == "dummy.html"

@patch('src.core_service.Path')
def test_convert_to_html_invalid_sheet_name(mock_path, core_service):
    """Test HTML conversion with an invalid sheet name."""
    mock_path.return_value.exists.return_value = True
    with pytest.raises(ValueError, match="工作表 'InvalidSheet' 在文件中未找到。"):
        core_service.convert_to_html("dummy.xlsx", sheet_name="InvalidSheet")


@patch('src.core_service.Path.exists', return_value=True)
def test_apply_changes_unsupported_file_type(mock_path_exists, core_service):
    """Test applying changes to an unsupported file type."""
    with pytest.raises(ValueError, match="Unsupported file type: .txt"):
        core_service.apply_changes("dummy.txt", {})


@patch('xlwt.Workbook')
def test_write_back_xls(mock_workbook, core_service):
    """Test writing back changes to an XLS file."""
    mock_sheet = MagicMock()
    mock_workbook.return_value.add_sheet.return_value = mock_sheet

    changes = {
        "sheet_name": "TestSheet",
        "headers": ["Header1", "Header2"],
        "rows": [
            [{"value": "R1C1"}, {"value": "R1C2"}],
            [{"value": "R2C1"}, {"value": "R2C2"}]
        ]
    }
    
    file_path = Path("dummy.xls")
    changes_applied = core_service._write_back_xls(file_path, changes)

    mock_workbook.return_value.add_sheet.assert_called_with("TestSheet")
    assert mock_sheet.write.call_count == 6  # 2 headers + 4 data cells
    mock_workbook.return_value.save.assert_called_with(str(file_path))
    assert changes_applied == 4


@patch("builtins.open", new_callable=mock_open)
def test_write_back_csv(mock_file, core_service):
    """Test writing back changes to a CSV file."""
    changes = {
        "sheet_name": "TestSheet",
        "headers": ["Header1", "Header2"],
        "rows": [
            [{"value": "R1C1"}, "R1C2_string"],
            [123, None]
        ]
    }
    
    file_path = Path("dummy.csv")
    changes_applied = core_service._write_back_csv(file_path, changes)

    # Verify the content that was written to the file
    mock_file().write.assert_any_call("Header1,Header2\r\n")
    mock_file().write.assert_any_call("R1C1,R1C2_string\r\n")
    mock_file().write.assert_any_call("123,\r\n")
    
    assert changes_applied == 2


@patch("builtins.open", new_callable=mock_open)
def test_write_back_csv_with_empty_row(mock_file, core_service):
    """Test writing back changes to a CSV file with an empty row."""
    changes = {
        "sheet_name": "TestSheet",
        "headers": ["Header1", "Header2"],
        "rows": [
            [{"value": "R1C1"}, {"value": "R1C2"}],
            [],
            [{"value": "R3C1"}, {"value": "R3C2"}]
        ]
    }
    
    file_path = Path("dummy.csv")
    changes_applied = core_service._write_back_csv(file_path, changes)

    # Verify the content that was written to the file
    # Note: The empty row should result in a line with just a newline character
    mock_file().write.assert_any_call("Header1,Header2\r\n")
    mock_file().write.assert_any_call("R1C1,R1C2\r\n")
    mock_file().write.assert_any_call("\r\n")
    mock_file().write.assert_any_call("R3C1,R3C2\r\n")
    
    assert changes_applied == 3


@patch('openpyxl.load_workbook')
def test_write_back_xlsx_sheet_not_found(mock_load_workbook, core_service):
    """Test _write_back_xlsx when the sheet name is not found in the workbook."""
    # 1. Create a mock workbook
    mock_wb = MagicMock()
    mock_wb.sheetnames = ['ExistingSheet']
    mock_load_workbook.return_value = mock_wb

    # 2. Define changes for a non-existing sheet
    changes = {
        "sheet_name": "NonExistingSheet",
        "headers": ["H1"],
        "rows": [[{"value": "v1"}]]
    }
    
    file_path = Path("dummy.xlsx")

    # 3. Assert that ValueError is raised
    with pytest.raises(ValueError, match="工作表 'NonExistingSheet' 在文件中不存在。"):
        core_service._write_back_xlsx(file_path, changes)

    mock_load_workbook.assert_called_once_with(file_path)
