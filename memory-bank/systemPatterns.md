# System Patterns: MCP Sheet Parser

## 1. 核心架构原则

- **最小可行架构**: 只实现必需的抽象，避免过度设计。
- **单一职责**: 每个模块、每个类、每个函数都只做一件事。
- **依赖注入**: 高层模块不依赖于低层模块的具体实现，两者都依赖于抽象。
- **组合优于继承**: 优先使用对象组合来实现代码复用。

## 2. 系统架构

### 2.1 目录结构 - 简化版
```
g:/MCP-Sheet-Parser/
├── src/
│   ├── main.py                 # MCP服务器入口（在根目录）
│   ├── core_service.py         # 核心业务服务（简化版）
│   ├── mcp_server/             # MCP服务器实现 (接口层)
│   │   ├── __init__.py
│   │   ├── server.py           # 服务器主逻辑
│   │   └── tools.py            # 3个核心工具定义
│   ├── parsers/                # 解析器 (实现层)
│   │   ├── __init__.py
│   │   ├── base_parser.py      # 抽象基类
│   │   ├── xlsx_parser.py      # .xlsx解析器
│   │   ├── csv_parser.py       # .csv解析器
│   │   └── factory.py          # 解析器工厂
│   ├── models/                 # 数据模型 (实现层)
│   │   ├── __init__.py
│   │   └── table_model.py      # Sheet/Row/Cell/Style 定义
│   └── exceptions/             # 自定义异常
│       ├── __init__.py
│       └── custom_exceptions.py
├── tests/
└── memory-bank/
```

**简化说明：**
- 移除了复杂的services、converters、utils、templates目录
- 用单一的core_service.py替代多层服务架构
- 专注于3个核心工具，避免过度工程化

### 2.2 类与接口定义

#### 2.2.1 `models/table_model.py`
```python
from dataclasses import dataclass, field
from typing import Any, List, Optional

@dataclass
class Style:
    # 字体属性
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_color: str = "#000000"
    font_size: float | None = None
    font_name: str | None = None

    # 背景和填充
    background_color: str = "#FFFFFF"

    # 文本对齐
    text_align: str = "left"  # left, center, right, justify
    vertical_align: str = "top"  # top, middle, bottom

    # 边框属性
    border_top: str = ""
    border_bottom: str = ""
    border_left: str = ""
    border_right: str = ""
    border_color: str = "#000000"

    # 文本换行和格式化
    wrap_text: bool = False
    number_format: str = ""

    # 进阶功能
    hyperlink: str | None = None  # 超链接URL
    comment: str | None = None    # 单元格注释

@dataclass
class Cell:
    value: Any
    style: Optional[Style] = None
    row_span: int = 1
    col_span: int = 1

@dataclass
class Row:
    cells: List[Cell]

@dataclass
class Sheet:
    name: str
    rows: List[Row]
    merged_cells: List[str] = field(default_factory=list)
```

#### 2.2.2 `parsers/base_parser.py`
```python
from abc import ABC, abstractmethod
from src.models.table_model import Sheet

class BaseParser(ABC):
    @abstractmethod
    def parse(self, file_path: str) -> Sheet:
        """Parses the given file and returns a Sheet object."""
        pass
```

#### 2.2.3 `parsers/parser_factory.py`
```python
from .base_parser import BaseParser
from .xlsx_parser import XlsxParser
from .csv_parser import CsvParser
from src.exceptions.custom_exceptions import UnsupportedFormatError

class ParserFactory:
    _parsers = {
        "csv": CsvParser(),
        "xlsx": XlsxParser(),
        "xls": XlsParser(),
        "xlsb": XlsbParser(),
    }

    @staticmethod
    def get_parser(file_path: str) -> BaseParser:
        file_extension = file_path.split('.')[-1].lower()
        parser = ParserFactory._parsers.get(file_extension)
        if not parser:
            supported_formats = ", ".join(ParserFactory._parsers.keys())
            raise UnsupportedFileType(
                f"不支持的文件格式: '{file_extension}'. "
                f"支持的格式: {supported_formats}"
            )
        return parser
```

#### 2.2.4 `converters/html_converter.py`
```python
from src.models.table_model import Sheet
# from jinja2 import Environment, FileSystemLoader

class HTMLConverter:
    def __init__(self):
        # self.env = Environment(loader=FileSystemLoader('src/templates'))
        pass

    def convert(self, sheet: Sheet) -> str:
        # template = self.env.get_template('table_template.html')
        # return template.render(sheet=sheet)
        # Placeholder implementation
        return "<html><body><table>...</table></body></html>"
```

#### 2.2.5 `services/sheet_service.py`
```python
from src.parsers.parser_factory import ParserFactory
from src.converters.html_converter import HTMLConverter

class SheetService:
    def __init__(self, parser_factory: ParserFactory, html_converter: HTMLConverter):
        self.parser_factory = parser_factory
        self.html_converter = html_converter

    def convert_to_html(self, file_path: str) -> str:
        file_extension = file_path.split('.')[-1]
        parser = self.parser_factory.get_parser(file_extension)
        sheet_model = parser.parse(file_path)
        html_content = self.html_converter.convert(sheet_model)
        return html_content
```

## 3. 数据流
`File Input` → `SheetService` → `ParserFactory` → `Specific Parser` → `Sheet Model` → `HTMLConverter` → `HTML Output`

## 4. MCP工具架构模式 - 简化版

### 4.1 三个核心工具设计
1. **`convert_to_html`** - 完美HTML转换
   - 直接文件到HTML，返回路径
   - 后台处理，不占用LLM上下文
   - 95%样式保真度

2. **`parse_sheet`** - JSON数据解析
   - 支持范围选择和工作表选择
   - 智能大小检测，大文件返回摘要
   - LLM友好的JSON格式

3. **`apply_changes`** - 数据写回
   - 完成编辑闭环
   - 自动备份机制
   - 保持文件格式和样式

### 4.2 工作流程模式
**HTML转换流程：**
```
文件 → convert_to_html → HTML文件路径
```

**数据分析流程：**
```
文件 → parse_sheet → JSON数据 → LLM分析
```

**数据编辑流程：**
```
文件 → parse_sheet → JSON → LLM编辑 → apply_changes → 写回文件
```

**大文件处理流程：**
```
文件 → parse_sheet → 摘要+建议 → parse_sheet(range) → 详细数据
```

### 4.3 JSON数据格式模式
```json
{
  "metadata": {"name": "sheet1", "rows": 100, "cols": 10},
  "data": [
    {"row": 0, "cells": [
      {"col": 0, "value": "Header1", "style": {"bold": true, "bg": "#f0f0f0"}},
      ...
    ]},
    ...
  ],
  "merged_cells": ["A1:B1", "C3:C5"],
  "styles": {"style_1": {"bold": true, "color": "#000"}}
}
```

## 5. 性能优化模式

### 5.1 智能处理策略
```python
class PerformanceOptimizer:
    SMALL_FILE_THRESHOLD = 1000      # 小文件：< 1000个单元格
    MEDIUM_FILE_THRESHOLD = 10000    # 中文件：< 10000个单元格
    LARGE_FILE_THRESHOLD = 50000     # 大文件：< 50000个单元格

    def get_processing_recommendation(self, total_cells: int) -> str:
        if total_cells < self.SMALL_FILE_THRESHOLD:
            return "small_file_direct"
        elif total_cells < self.MEDIUM_FILE_THRESHOLD:
            return "medium_file_direct"
        elif total_cells < self.LARGE_FILE_THRESHOLD:
            return "large_file_recommend_file_output"
        else:
            return "very_large_file_require_pagination"
```

### 5.2 分页处理模式
```python
def calculate_pagination_params(self, sheet: Sheet, max_rows_per_page: int = 1000):
    total_rows = len(sheet.rows)
    total_pages = (total_rows + max_rows_per_page - 1) // max_rows_per_page

    return {
        "total_rows": total_rows,
        "total_pages": total_pages,
        "pages": [
            {
                "page_number": i + 1,
                "start_row": i * max_rows_per_page,
                "end_row": min((i + 1) * max_rows_per_page, total_rows),
                "row_count": min(max_rows_per_page, total_rows - i * max_rows_per_page)
            }
            for i in range(total_pages)
        ]
    }
```

## 6. 错误处理
- 自定义异常（如 `UnsupportedFormatError`, `ParsingError`）在 `exceptions/custom_exceptions.py` 中定义。
- `SheetService` 捕获并向上抛出这些异常，由 `mcp_server` 统一处理并返回给用户。