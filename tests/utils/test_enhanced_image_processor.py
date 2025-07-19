import pytest
from src.utils.enhanced_image_processor import EnhancedImageProcessor

@pytest.fixture
def processor():
    """提供一个 EnhancedImageProcessor 实例。"""
    return EnhancedImageProcessor()

# 使用正确的字节序列创建用于测试的伪造图片数据
FAKE_PNG = b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR'
FAKE_JPEG = b'\xff\xd8\xff\xe0\x00\x10JFIF'
FAKE_GIF = b'GIF89a\x01\x00\x01\x00'
FAKE_BMP = b'BM\x36\x00\x00\x00'
INVALID_DATA = b'this is not an image'
SHORT_DATA = b'\x89PNG'

def test_detect_image_format(processor):
    """测试 detect_image_format 方法。"""
    assert processor.detect_image_format(FAKE_PNG) == 'png'
    assert processor.detect_image_format(FAKE_JPEG) == 'jpeg'
    assert processor.detect_image_format(FAKE_GIF) == 'gif'
    assert processor.detect_image_format(FAKE_BMP) == 'bmp'
    assert processor.detect_image_format(INVALID_DATA) == 'unknown'
    assert processor.detect_image_format(b'') == 'unknown'

def test_validate_image_data(processor):
    """测试 validate_image_data 方法。"""
    assert processor.validate_image_data(FAKE_PNG) is True
    assert processor.validate_image_data(INVALID_DATA) is False
    assert processor.validate_image_data(SHORT_DATA) is False
    assert processor.validate_image_data(b'') is False

def test_optimize_image_size(processor):
    """测试 optimize_image_size 方法（当前为占位符）。"""
    assert processor.optimize_image_size(FAKE_PNG, max_size=10) == FAKE_PNG
    assert processor.optimize_image_size(FAKE_PNG, max_size=1000) == FAKE_PNG

def test_generate_image_html_valid(processor):
    """测试使用有效数据生成HTML。"""
    html = processor.generate_image_html(FAKE_PNG, alt_text="A fake PNG")
    assert '<img src="data:image/png;base64,' in html
    assert 'alt="A fake PNG"' in html
    assert '图片加载失败' not in html

def test_generate_image_html_invalid(processor):
    """测试使用无效数据生成占位符HTML。"""
    html = processor.generate_image_html(INVALID_DATA, alt_text="Invalid Image")
    assert '图片加载失败: Invalid Image' in html
    assert '<img' not in html

def test_generate_placeholder_html(processor):
    """测试 _generate_placeholder_html 方法。"""
    html = processor._generate_placeholder_html("Placeholder")
    assert '图片加载失败: Placeholder' in html