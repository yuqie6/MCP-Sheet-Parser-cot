import pytest
from abc import ABC
from src.parsers.base_parser import BaseParser
from src.models.table_model import Sheet, LazySheet


class ConcreteParser(BaseParser):
    """具体的解析器实现，用于测试抽象基类"""
    
    def parse(self, file_path: str) -> list[Sheet]:
        """实现抽象方法"""
        return [Sheet(name="TestSheet", rows=[])]

class StreamingParser(BaseParser):
    """支持流式处理的解析器实现"""
    
    def parse(self, file_path: str) -> list[Sheet]:
        """实现抽象方法"""
        return [Sheet(name="StreamingSheet", rows=[])]
    
    def supports_streaming(self) -> bool:
        """
        TDD测试：supports_streaming应该能被子类重写
        
        这个测试覆盖第42-43行的代码路径
        """
        return True
    
    def create_lazy_sheet(self, file_path: str, sheet_name: str | None = None) -> LazySheet | None:
        """
        TDD测试：create_lazy_sheet应该能被子类重写

        这个测试覆盖第45-59行的代码路径
        """
        from unittest.mock import MagicMock
        mock_provider = MagicMock()
        return LazySheet(name=sheet_name or "default", provider=mock_provider)

def test_base_parser_is_abstract():
    """
    TDD测试：BaseParser应该是抽象类，不能直接实例化
    
    这个测试确保BaseParser是正确的抽象基类
    """
    # 验证BaseParser是ABC的子类
    assert issubclass(BaseParser, ABC)
    
    # 尝试直接实例化应该失败
    with pytest.raises(TypeError):
        BaseParser()

def test_concrete_parser_can_be_instantiated():
    """
    TDD测试：具体的解析器实现应该能够被实例化
    
    这个测试验证抽象方法的实现
    """
    parser = ConcreteParser()
    # 验证可以调用parse方法
    sheets = parser.parse("dummy.txt")
    assert len(sheets) == 1
    assert sheets[0].name == "TestSheet"

def test_default_supports_streaming():
    """
    TDD测试：BaseParser的默认supports_streaming应该返回False
    
    这个测试覆盖第41-43行的代码路径
    """
    parser = ConcreteParser()
    
    # 默认应该不支持流式处理
    assert parser.supports_streaming() is False

def test_default_create_lazy_sheet():
    """
    TDD测试：BaseParser的默认create_lazy_sheet应该返回None
    
    这个测试覆盖第45-59行的代码路径
    """
    parser = ConcreteParser()
    
    # 默认应该返回None
    result = parser.create_lazy_sheet("dummy.txt")
    assert result is None
    
    # 测试带sheet_name参数的情况
    result = parser.create_lazy_sheet("dummy.txt", "Sheet1")
    assert result is None

def test_streaming_parser_overrides():
    """
    TDD测试：子类应该能够重写supports_streaming和create_lazy_sheet
    
    这个测试验证子类可以正确重写基类方法
    """
    parser = StreamingParser()
    # 验证重写的supports_streaming
    assert parser.supports_streaming() is True
    
    # 验证重写的create_lazy_sheet
    lazy_sheet = parser.create_lazy_sheet("test.xlsx")
    assert lazy_sheet is not None
    assert isinstance(lazy_sheet, LazySheet)
    assert lazy_sheet.name == "default"
    
    # 测试带sheet_name参数的情况
    lazy_sheet = parser.create_lazy_sheet("test.xlsx", "CustomSheet")
    assert lazy_sheet.name == "CustomSheet"

def test_parse_method_is_abstract():
    """
    TDD测试：parse方法应该是抽象的，必须被子类实现
    
    这个测试验证抽象方法的强制实现
    """
    # 创建一个不实现parse方法的类应该失败
    with pytest.raises(TypeError):
        class IncompleteParser(BaseParser):
            pass
        IncompleteParser()

def test_base_parser_interface_completeness():
    """
    TDD测试：BaseParser应该定义完整的接口
    
    这个测试验证基类定义了所有必要的方法
    """
    # 验证BaseParser有所有预期的方法
    assert hasattr(BaseParser, 'parse')
    assert hasattr(BaseParser, 'supports_streaming')
    assert hasattr(BaseParser, 'create_lazy_sheet')
    
    # 验证parse是抽象方法
    assert getattr(BaseParser.parse, '__isabstractmethod__', False)
    
    # 验证其他方法不是抽象的（有默认实现）
    assert not getattr(BaseParser.supports_streaming, '__isabstractmethod__', False)
    assert not getattr(BaseParser.create_lazy_sheet, '__isabstractmethod__', False)

class CustomBehaviorParser(BaseParser):
    """测试自定义行为的解析器"""
    
    def __init__(self, should_support_streaming=False, lazy_sheet_result=None):
        self.should_support_streaming = should_support_streaming
        self.lazy_sheet_result = lazy_sheet_result
    
    def parse(self, file_path: str) -> list[Sheet]:
        return [Sheet(name=f"Sheet from {file_path}", rows=[])]
    
    def supports_streaming(self) -> bool:
        return self.should_support_streaming
    
    def create_lazy_sheet(self, file_path: str, sheet_name: str | None = None) -> LazySheet | None:
        return self.lazy_sheet_result

def test_parser_with_custom_streaming_behavior():
    """
    TDD测试：解析器应该能够自定义流式处理行为
    
    这个测试验证子类可以灵活地实现不同的行为
    """
    # 测试不支持流式处理的解析器
    parser1 = CustomBehaviorParser(should_support_streaming=False)
    assert parser1.supports_streaming() is False

    # 测试支持流式处理的解析器
    parser2 = CustomBehaviorParser(should_support_streaming=True)
    assert parser2.supports_streaming() is True

def test_parser_with_custom_lazy_sheet_behavior():
    """
    TDD测试：解析器应该能够自定义LazySheet创建行为
    
    这个测试验证子类可以返回不同类型的LazySheet结果
    """
    # 测试返回None的解析器
    parser1 = CustomBehaviorParser(lazy_sheet_result=None)
    assert parser1.create_lazy_sheet("test.xlsx") is None

    # 测试返回LazySheet的解析器
    from unittest.mock import MagicMock
    mock_provider = MagicMock()
    lazy_sheet = LazySheet(name="TestSheet", provider=mock_provider)
    parser2 = CustomBehaviorParser(lazy_sheet_result=lazy_sheet)
    result = parser2.create_lazy_sheet("test.xlsx")
    assert result is lazy_sheet
    assert result.name == "TestSheet"


class IncompleteParser(BaseParser):
    """不完整的解析器实现，用于测试抽象方法的直接调用"""

    # 故意不实现parse方法，保持抽象状态
    pass

def test_abstract_parse_method_direct_call():
    """
    TDD测试：抽象parse方法的pass语句应该被覆盖

    这个测试覆盖第39行的pass语句
    """
    # 直接调用BaseParser的parse方法（绕过抽象检查）
    result = BaseParser.parse(None, "dummy_path")

    # 抽象方法的pass语句应该返回None
    assert result is None
