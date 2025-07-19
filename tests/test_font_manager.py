
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

# === TDD测试：提升FontManager覆盖率到100% ===

@patch("builtins.open")
@patch("pathlib.Path.exists", return_value=True)
def test_font_manager_init_with_json_decode_error(mock_exists, mock_file):
    """
    TDD测试：FontManager应该处理JSON解码错误

    这个测试覆盖第47-49行的JSON解码错误处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    # 模拟无效的JSON内容
    mock_file.return_value = mock_open(read_data="invalid json content").return_value

    # 应该使用默认配置而不是崩溃
    fm = FontManager(config_file="invalid.json")
    assert "SimSun" in fm.font_database["chinese_keywords"]
    assert isinstance(fm.custom_mappings, dict)

@patch("builtins.open", side_effect=IOError("File read error"))
@patch("pathlib.Path.exists", return_value=True)
def test_font_manager_init_with_file_read_error(mock_exists, mock_file):
    """
    TDD测试：FontManager应该处理文件读取错误

    这个测试覆盖第50-52行的文件读取错误处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为

    # 应该使用默认配置而不是崩溃
    fm = FontManager(config_file="unreadable.json")
    assert "SimSun" in fm.font_database["chinese_keywords"]
    assert isinstance(fm.custom_mappings, dict)

def test_detect_font_type_with_custom_mapping():
    """
    TDD测试：detect_font_type应该使用自定义映射

    这个测试覆盖第65-66行的自定义映射处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()
    fm.custom_mappings = {"CustomFont": "Arial"}

    # 应该使用映射后的字体进行检测
    font_type = fm.detect_font_type("CustomFont")
    assert font_type == "sans_serif"  # Arial的类型

def test_detect_font_type_with_unknown_font():
    """
    TDD测试：detect_font_type应该处理未知字体

    这个测试覆盖第85行的默认返回代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()

    # 未知字体应该返回默认类型
    font_type = fm.detect_font_type("UnknownFont")
    assert font_type == "sans_serif"

def test_format_font_name_with_special_characters():
    """
    TDD测试：format_font_name应该处理特殊字符

    这个测试确保包含特殊字符的字体名被正确格式化
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()

    # 包含特殊字符的字体名应该被引号包围
    assert fm.format_font_name("Font-Name") == '"Font-Name"'
    assert fm.format_font_name("Font_Name") == '"Font_Name"'
    assert fm.format_font_name("Font Name 123") == '"Font Name 123"'

def test_generate_font_family_with_chinese_font():
    """
    TDD测试：generate_font_family应该为中文字体添加后备字体

    这个测试覆盖第103-104行的中文字体处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()

    font_family = fm.generate_font_family("SimSun")

    # 应该包含中文后备字体
    assert "SimSun" in font_family
    assert "sans-serif" in font_family

def test_generate_font_family_with_monospace_font():
    """
    TDD测试：generate_font_family应该为等宽字体添加后备字体

    这个测试覆盖第105-106行的等宽字体处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()

    font_family = fm.generate_font_family("Courier New")

    # 应该包含等宽后备字体
    assert "Courier New" in font_family
    assert "monospace" in font_family

def test_generate_font_family_with_serif_font():
    """
    TDD测试：generate_font_family应该为衬线字体添加后备字体

    这个测试覆盖第107-108行的衬线字体处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()

    font_family = fm.generate_font_family("Times New Roman")

    # 应该包含衬线后备字体
    assert "Times New Roman" in font_family
    assert "serif" in font_family

def test_generate_font_family_with_sans_serif_font():
    """
    TDD测试：generate_font_family应该为无衬线字体添加后备字体

    这个测试覆盖第109-110行的无衬线字体处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()

    font_family = fm.generate_font_family("Arial")

    # 应该包含无衬线后备字体
    assert "Arial" in font_family
    assert "sans-serif" in font_family

def test_get_font_info_with_custom_mapped_font():
    """
    TDD测试：get_font_info应该处理自定义映射的字体

    这个测试确保自定义映射在get_font_info中正确工作
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()
    fm.custom_mappings = {"MyCustomFont": "Arial"}

    info = fm.get_font_info("MyCustomFont")

    # 应该返回映射后字体的信息
    assert info['font_type'] == 'sans_serif'
    assert info['needs_quotes'] is False
    assert info['formatted_name'] == 'MyCustomFont'

def test_get_font_info_with_font_needing_quotes():
    """
    TDD测试：get_font_info应该正确标识需要引号的字体

    这个测试确保needs_quotes字段被正确设置
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()

    info = fm.get_font_info("Font With Spaces")

    # 包含空格的字体名应该需要引号
    assert info['needs_quotes'] is True
    assert info['formatted_name'] == '"Font With Spaces"'

def test_get_font_manager_singleton():
    """
    TDD测试：get_font_manager应该返回单例实例

    这个测试覆盖第125行的单例模式代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为

    # 多次调用应该返回同一个实例
    fm1 = get_font_manager()
    fm2 = get_font_manager()

    assert fm1 is fm2
    assert isinstance(fm1, FontManager)

def test_font_database_completeness():
    """
    TDD测试：验证字体数据库包含所有预期的字体类别

    这个测试确保默认字体数据库的完整性
    """
    # 🔴 红阶段：编写测试描述期望的行为
    fm = FontManager()

    # 验证所有字体类别都存在
    expected_categories = ["chinese_keywords", "monospace_keywords", "serif_keywords", "sans_serif_keywords"]

    for category in expected_categories:
        assert category in fm.font_database
        assert isinstance(fm.font_database[category], list)
        assert len(fm.font_database[category]) > 0
