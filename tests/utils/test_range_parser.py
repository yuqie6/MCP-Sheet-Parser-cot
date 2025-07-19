import pytest
from src.utils.range_parser import parse_range_string

def test_parse_single_cell():
    """测试解析单个单元格字符串。"""
    assert parse_range_string("A1") == (0, 0, 0, 0)
    assert parse_range_string("C10") == (9, 2, 9, 2)
    assert parse_range_string("Z100") == (99, 25, 99, 25)

def test_parse_multi_letter_column_single_cell():
    """测试解析多字母列的单个单元格。"""
    assert parse_range_string("AA1") == (0, 26, 0, 26)
    assert parse_range_string("BC22") == (21, 54, 21, 54)

def test_parse_range():
    """测试解析单元格范围字符串。"""
    assert parse_range_string("A1:D10") == (0, 0, 9, 3)
    assert parse_range_string("C5:F8") == (4, 2, 7, 5)
    assert parse_range_string("Y80:AC95") == (79, 24, 94, 28)

def test_parse_case_insensitivity():
    """测试解析时的大小写不敏感性。"""
    assert parse_range_string("a1:d10") == (0, 0, 9, 3)
    assert parse_range_string("c5") == (4, 2, 4, 2)

def test_parse_with_whitespace():
    """测试解析时处理前后空格。"""
    assert parse_range_string("  A1:D10  ") == (0, 0, 9, 3)
    assert parse_range_string("  B2  ") == (1, 1, 1, 1)

def test_invalid_range_format():
    """测试无效的范围格式是否会引发ValueError。"""
    with pytest.raises(ValueError, match="无效的范围格式: A1:B"):
        parse_range_string("A1:B")
    
    with pytest.raises(ValueError, match="无效的范围格式: INVALID"):
        parse_range_string("INVALID")

    with pytest.raises(ValueError, match="无效的范围格式: 1:10"):
        parse_range_string("1:10")