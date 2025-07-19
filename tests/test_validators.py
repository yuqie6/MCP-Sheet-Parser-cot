
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
