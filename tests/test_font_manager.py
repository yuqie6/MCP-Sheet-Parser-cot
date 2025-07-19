
import pytest
import json
from unittest.mock import patch, mock_open
from src.font_manager import FontManager, get_font_manager

@pytest.fixture
def mock_config_file():
    """Fixture for a mocked font config file."""
    config_data = {
        "font_database": {
            "chinese_keywords": ["TestFont"],
        },
        "custom_mappings": {
            "MyFont": "YourFont"
        }
    }
    return json.dumps(config_data)

@patch("builtins.open", new_callable=mock_open)
@patch("pathlib.Path.exists", return_value=False)
def test_font_manager_init_defaults(mock_exists, mock_file):
    """Test FontManager initialization with default settings."""
    fm = FontManager()
    assert "SimSun" in fm.font_database["chinese_keywords"]

@patch("builtins.open")
@patch("pathlib.Path.exists", return_value=True)
def test_font_manager_init_with_config(mock_exists, mock_file, mock_config_file):
    """Test FontManager initialization with a custom config file."""
    mock_file.return_value = mock_open(read_data=mock_config_file).return_value
    fm = FontManager(config_file="dummy_config.json")
    assert "TestFont" in fm.font_database["chinese_keywords"]
    assert fm.custom_mappings["MyFont"] == "YourFont"

def test_detect_font_type():
    """Test font type detection."""
    fm = FontManager()
    assert fm.detect_font_type("SimSun") == "chinese"
    assert fm.detect_font_type("Courier New") == "monospace"
    assert fm.detect_font_type("Times New Roman") == "serif"
    assert fm.detect_font_type("Arial") == "sans_serif"

def test_format_font_name():
    """Test font name formatting."""
    fm = FontManager()
    assert fm.format_font_name("My Font") == '"My Font"'
    assert fm.format_font_name("Arial") == "Arial"

def test_generate_font_family():
    """Test font-family string generation."""
    fm = FontManager()
    font_family = fm.generate_font_family("SimSun")
    assert "SimSun" in font_family
    assert "Microsoft YaHei" in font_family

def test_get_font_manager_singleton():
    """Test that get_font_manager returns a singleton."""
    fm1 = get_font_manager()
    fm2 = get_font_manager()
    assert fm1 is fm2

def test_learn_font():
    """Test learning a new font."""
    fm = FontManager()
    fm.learn_font("My New Font", "sans_serif")
    assert "My" in fm.font_database["sans_serif_keywords"]
    assert "New" in fm.font_database["sans_serif_keywords"]
    assert "Font" in fm.font_database["sans_serif_keywords"]

@patch("builtins.open", new_callable=mock_open)
@patch("pathlib.Path.exists", return_value=True)
def test_save_config(mock_exists, mock_file):
    """Test saving the font config."""
    fm = FontManager()
    fm.add_custom_mapping("NewMap", "MappedFont")
    fm.save_config()
    mock_file().write.assert_called()

def test_get_font_info():
    """Test getting font info."""
    fm = FontManager()
    info = fm.get_font_info("Arial")
    assert info['font_type'] == 'sans_serif'
    assert info['needs_quotes'] is False
