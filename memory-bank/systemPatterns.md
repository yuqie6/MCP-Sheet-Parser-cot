# System Patterns: MCP Sheet Parser

## 1. 核心架构原则

- **最小可行架构**: 只实现必需的抽象，避免过度设计。
- **单一职责**: 每个模块、每个类、每个函数都只做一件事。
- **依赖注入**: 高层模块不依赖于低层模块的具体实现，两者都依赖于抽象。
- **组合优于继承**: 优先使用对象组合来实现代码复用。

## 2. 系统架构

### 2.1 目录结构
```
g:/MCP-Sheet-Parser/
├── src/
│   ├── main.py                 # MCP服务器入口
│   ├── mcp_server/             # MCP服务器实现 (接口层)
│   │   ├── __init__.py
│   │   └── tools.py            # 工具定义
│   ├── services/               # 业务逻辑层
│   │   ├── __init__.py
│   │   └── sheet_service.py    # 核心业务服务
│   ├── parsers/                # 解析器 (实现层)
│   │   ├── __init__.py
│   │   ├── base_parser.py      # 抽象基类
│   │   ├── xlsx_parser.py      # .xlsx解析器
│   │   ├── csv_parser.py       # .csv解析器
│   │   └── parser_factory.py   # 解析器工厂
│   ├── converters/             # 转换器 (实现层)
│   │   ├── __init__.py
│   │   └── html_converter.py   # HTML转换器
│   ├── models/                 # 数据模型 (实现层)
│   │   ├── __init__.py
│   │   └── table_model.py      # Sheet/Row/Cell/Style 定义
│   ├── templates/              # HTML模板
│   │   └── table_template.html
│   └── exceptions/             # 自定义异常
│       ├── __init__.py
│       └── custom_exceptions.py
├── tests/
└── memory-bank/
```

### 2.2 类与接口定义

#### 2.2.1 `models/table_model.py`
```python
from dataclasses import dataclass, field
from typing import Any, List, Optional

@dataclass
class Style:
    bold: bool = False
    italic: bool = False
    font_color: str = "#000000"
    background_color: str = "#FFFFFF"

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

## 4. MCP工具架构模式

### 4.1 工具分层设计
**核心工具（90%使用场景）：**
- `parse_sheet_to_json` - 数据获取和分析
- `convert_json_to_html` - 完美复刻生成
- `convert_file_to_html` - 智能直接转换

**专业工具（高级需求）：**
- `convert_file_to_html_file` - 文件输出模式
- `get_table_summary` - 预览分析
- `get_sheet_metadata` - 元数据获取

### 4.2 工作流程模式
**完美复刻流程：**
```
文件 → parse_sheet_to_json → LLM分析 → convert_json_to_html → 完美HTML文件
```

**快速转换流程：**
```
文件 → convert_file_to_html → 智能HTML输出
```

**预览分析流程：**
```
文件 → get_table_summary → 摘要信息
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

## 5. 错误处理
- 自定义异常（如 `UnsupportedFormatError`, `ParsingError`）在 `exceptions/custom_exceptions.py` 中定义。
- `SheetService` 捕获并向上抛出这些异常，由 `mcp_server` 统一处理并返回给用户。