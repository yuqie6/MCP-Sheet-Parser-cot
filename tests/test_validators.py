
import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path
from src.validators import FileValidator, RangeValidator, DataValidator, validate_file_input
from src.exceptions import ValidationError, FileNotFoundError, FileAccessError, UnsupportedFileTypeError

# --- FileValidator Tests ---

@patch('src.validators.Path.exists', return_value=True)
@patch('src.validators.Path.is_file', return_value=True)
@patch('os.access', return_value=True)
def test_file_validator_path_success(mock_access, mock_is_file, mock_exists):
    """Test successful file path validation."""
    with patch('src.validators.Path.resolve', return_value=Path("/test/dummy.xlsx")):
        path = FileValidator.validate_file_path("/test/dummy.xlsx")
        assert isinstance(path, Path)

def test_file_validator_path_non_existent():
    """Test file path that does not exist."""
    with patch('src.validators.Path.resolve', return_value=Path("/nonexistent/dummy.xlsx")):
        with pytest.raises(FileNotFoundError):
            FileValidator.validate_file_path("/nonexistent/dummy.xlsx")

def test_file_validator_dangerous_path():
    """Test dangerous file path patterns."""
    with pytest.raises(ValidationError):
        FileValidator.validate_file_path("../../etc/passwd")

def test_file_validator_extension_supported():
    """Test supported file extensions."""
    assert FileValidator.validate_file_extension("test.csv") == "csv"
    assert FileValidator.validate_file_extension("test.XLSX") == "xlsx"

def test_file_validator_extension_unsupported():
    """Test unsupported file extensions."""
    with pytest.raises(UnsupportedFileTypeError):
        FileValidator.validate_file_extension("test.txt")

@patch('src.validators.Path.stat')
@patch('src.validators.FileValidator.get_max_file_size_mb', return_value=10)
def test_file_validator_size(mock_max_size, mock_stat):
    """Test file size validation."""
    mock_stat.return_value.st_size = 5 * 1024 * 1024 # 5MB
    assert FileValidator.validate_file_size("small.xlsx") > 0

    mock_stat.return_value.st_size = 15 * 1024 * 1024 # 15MB
    with pytest.raises(ValidationError):
        FileValidator.validate_file_size("large.xlsx")

# --- RangeValidator Tests ---

def test_range_validator_valid():
    """Test valid range strings."""
    assert RangeValidator.validate_range_string("A1:B10") == ("A1", "B10")
    assert RangeValidator.validate_range_string("C1") == ("C1", "C1")

def test_range_validator_invalid():
    """Test invalid range strings."""
    with pytest.raises(ValidationError):
        RangeValidator.validate_range_string("A1:B")
    with pytest.raises(ValidationError):
        RangeValidator.validate_range_string("1A:2B")

# --- DataValidator Tests ---

def test_data_validator_sheet_name():
    """Test sheet name validation."""
    assert DataValidator.validate_sheet_name("Valid Name") == "Valid Name"
    with pytest.raises(ValidationError):
        DataValidator.validate_sheet_name("Invalid[Name]")
    with pytest.raises(ValidationError):
        DataValidator.validate_sheet_name("a" * 32)

# --- Combined Validator Test ---

@patch('src.validators.FileValidator.validate_file_path')
@patch('src.validators.FileValidator.validate_file_extension')
@patch('src.validators.FileValidator.validate_file_size')
def test_validate_file_input(mock_size, mock_ext, mock_path):
    """Test the main validate_file_input function."""
    mock_path.return_value = Path("valid.xlsx")
    mock_ext.return_value = "xlsx"
    mock_size.return_value = 1024
    path, ext = validate_file_input("valid.xlsx")
    assert path == Path("valid.xlsx")
    assert ext == "xlsx"

# --- Additional Tests for Better Coverage ---

def test_file_validator_empty_path():
    """Test validation with empty file path."""
    with pytest.raises(ValidationError, match="文件路径不能为空且必须是字符串"):
        FileValidator.validate_file_path("")

    with pytest.raises(ValidationError, match="文件路径不能为空且必须是字符串"):
        FileValidator.validate_file_path(None)

def test_file_validator_non_string_path():
    """Test validation with non-string file path."""
    with pytest.raises(ValidationError, match="文件路径不能为空且必须是字符串"):
        FileValidator.validate_file_path(123)

def test_file_validator_invalid_path_format():
    """Test validation with invalid path format."""
    with patch('pathlib.Path.resolve', side_effect=OSError("Invalid path")):
        with pytest.raises(ValidationError, match="路径格式无效"):
            FileValidator.validate_file_path("invalid\x00path")

@patch('src.validators.Path.exists', return_value=True)
@patch('src.validators.Path.is_file', return_value=False)
def test_file_validator_path_is_directory(mock_is_file, mock_exists):
    """Test validation when path points to a directory."""
    with patch('src.validators.Path.resolve', return_value=Path("/test/directory")):
        with pytest.raises(ValidationError, match="路径必须指向一个文件，不能是目录"):
            FileValidator.validate_file_path("/test/directory")

@patch('src.validators.Path.exists', return_value=True)
@patch('src.validators.Path.is_file', return_value=True)
@patch('os.access', return_value=False)
def test_file_validator_no_read_permission(mock_access, mock_is_file, mock_exists):
    """Test validation when file has no read permission."""
    with patch('src.validators.Path.resolve', return_value=Path("/test/noperm.xlsx")):
        with pytest.raises(FileAccessError):
            FileValidator.validate_file_path("/test/noperm.xlsx")

def test_file_validator_windows_path_traversal():
    """Test Windows path traversal detection."""
    with pytest.raises(ValidationError, match="检测到危险路径模式"):
        FileValidator.validate_file_path("..\\..\\windows\\system32\\config")

def test_file_validator_unix_system_path():
    """Test Unix system path detection."""
    with pytest.raises(ValidationError, match="检测到危险路径模式"):
        FileValidator.validate_file_path("/etc/passwd")

def test_file_validator_get_max_file_size():
    """Test getting max file size from config."""
    with patch('src.unified_config.get_config') as mock_get_config:
        mock_config = MagicMock()
        mock_config.max_file_size_mb = 50
        mock_get_config.return_value = mock_config

        max_size = FileValidator.get_max_file_size_mb()
        assert max_size == 50

def test_range_validator_invalid_format():
    """Test range validation with invalid format."""
    with pytest.raises(ValidationError, match="范围格式无效"):
        RangeValidator.validate_range_string("invalid_range")

def test_range_validator_invalid_cell_reference():
    """Test range validation with valid cell reference."""
    result = RangeValidator.validate_range_string("A1:B10")
    assert result == ("A1", "B10")

    result = RangeValidator.validate_range_string("A1")
    assert result == ("A1", "A1")

def test_data_validator_invalid_sheet_name():
    """Test sheet name validation with invalid characters."""
    with pytest.raises(ValidationError, match="工作表名称不能包含字符"):
        DataValidator.validate_sheet_name("Sheet[invalid]")

    with pytest.raises(ValidationError, match="工作表名称不能包含字符"):
        DataValidator.validate_sheet_name("Sheet*name")

def test_data_validator_empty_sheet_name():
    """Test sheet name validation with empty name."""
    with pytest.raises(ValidationError, match="工作表名称不能为空"):
        DataValidator.validate_sheet_name("")

def test_data_validator_whitespace_sheet_name():
    """Test sheet name validation with whitespace name."""
    # 空白字符串会被strip()处理，如果结果为空则抛出异常
    result = DataValidator.validate_sheet_name("   valid   ")
    assert result == "valid"

def test_data_validator_long_sheet_name():
    """Test sheet name validation with too long name."""
    long_name = "a" * 32  # Excel限制是31个字符
    with pytest.raises(ValidationError, match="工作表名称不能超过31个字符"):
        DataValidator.validate_sheet_name(long_name)

def test_data_validator_validate_page_size():
    """Test page size validation."""
    result = DataValidator.validate_page_size(10)
    assert result == 10

    with pytest.raises(ValidationError, match="分页大小必须是整数"):
        DataValidator.validate_page_size("10")

    with pytest.raises(ValidationError, match="分页大小必须大于0"):
        DataValidator.validate_page_size(0)

    with pytest.raises(ValidationError, match="分页大小不能超过10000"):
        DataValidator.validate_page_size(10001)

def test_range_validator_empty_range():
    """Test range validation with empty range."""
    result = RangeValidator.validate_range_string("")
    assert result is None

    result = RangeValidator.validate_range_string(None)
    assert result is None

def test_range_validator_non_string_range():
    """Test range validation with non-string range."""
    with pytest.raises(ValidationError, match="范围必须是字符串"):
        RangeValidator.validate_range_string(123)

def test_data_validator_non_string_sheet_name():
    """Test sheet name validation with non-string name."""
    with pytest.raises(ValidationError, match="工作表名称必须是字符串"):
        DataValidator.validate_sheet_name(123)

def test_data_validator_validate_page_number():
    """Test page number validation."""
    result = DataValidator.validate_page_number(1)
    assert result == 1

    with pytest.raises(ValidationError, match="页码必须是整数"):
        DataValidator.validate_page_number("1")

    with pytest.raises(ValidationError, match="页码必须大于0"):
        DataValidator.validate_page_number(0)

def test_data_validator_validate_output_path():
    """Test output path validation."""
    with patch('pathlib.Path.resolve') as mock_resolve:
        with patch('pathlib.Path.parent') as mock_parent:
            with patch('pathlib.Path.exists', return_value=True):
                with patch('os.access', return_value=True):
                    mock_path = MagicMock()
                    mock_resolve.return_value = mock_path
                    mock_parent.return_value = mock_path

                    result = DataValidator.validate_output_path("/test/output.txt")
                    assert result == mock_path

def test_data_validator_validate_output_path_empty():
    """Test output path validation with empty path."""
    with pytest.raises(ValidationError, match="输出路径不能为空且必须是字符串"):
        DataValidator.validate_output_path("")

    with pytest.raises(ValidationError, match="输出路径不能为空且必须是字符串"):
        DataValidator.validate_output_path(None)

def test_data_validator_validate_output_path_invalid_format():
    """Test output path validation with invalid format."""
    with patch('pathlib.Path.resolve', side_effect=OSError("Invalid path")):
        with pytest.raises(ValidationError, match="路径格式无效"):
            DataValidator.validate_output_path("invalid\x00path")

def test_data_validator_validate_output_path_create_dirs(tmp_path):
    """Test output path validation with directory creation."""
    # 使用临时目录进行真实测试
    output_file = tmp_path / "new" / "output.txt"

    result = DataValidator.validate_output_path(str(output_file), create_dirs=True)

    # 验证目录被创建
    assert output_file.parent.exists()
    assert result == output_file

def test_data_validator_validate_output_path_no_create_dirs():
    """Test output path validation without directory creation."""
    with patch('pathlib.Path') as mock_path_class:
        mock_path = MagicMock()
        mock_parent = MagicMock()
        mock_parent.exists.return_value = False
        mock_path.parent = mock_parent
        mock_path_class.return_value.resolve.return_value = mock_path

        with pytest.raises(ValidationError, match="目录不存在"):
            DataValidator.validate_output_path("/test/new/output.txt", create_dirs=False)

def test_data_validator_validate_output_path_mkdir_error():
    """Test output path validation with mkdir error."""
    with patch('pathlib.Path') as mock_path_class:
        mock_path = MagicMock()
        mock_parent = MagicMock()
        mock_parent.exists.return_value = False
        mock_parent.mkdir.side_effect = OSError("Permission denied")
        mock_path.parent = mock_parent
        mock_path_class.return_value.resolve.return_value = mock_path

        with pytest.raises(ValidationError, match="无法创建目录"):
            DataValidator.validate_output_path("/test/new/output.txt", create_dirs=True)

def test_data_validator_validate_output_path_no_write_permission():
    """Test output path validation with no write permission."""
    with patch('pathlib.Path.resolve') as mock_resolve:
        with patch('pathlib.Path.parent') as mock_parent:
            with patch('pathlib.Path.exists', return_value=True):
                with patch('os.access', return_value=False):
                    mock_path = MagicMock()
                    mock_resolve.return_value = mock_path
                    mock_parent.return_value = mock_path

                    with pytest.raises(FileAccessError):
                        DataValidator.validate_output_path("/test/output.txt")
