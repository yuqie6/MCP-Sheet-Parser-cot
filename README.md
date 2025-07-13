# MCP 表格解析器

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io/)

一个专为 **模型上下文协议 (MCP) 服务器** 设计的高保真表格解析器和HTML转换器。该工具使llm能够无缝读取、分析和修改表格文件、转化为html，并完整保留样式。

## 🚀 什么是 MCP 表格解析器？

MCP 表格解析器是一个 **模型上下文协议服务器**，为AI助手提供强大的表格处理能力。它充当AI模型与表格文件之间的桥梁，提供三个核心工具：

- **`parse_sheet`** - 将任何表格解析为AI友好的JSON格式
- **`convert_to_html`** - 生成保留样式的高保真HTML
- **`apply_changes`** - 将修改后的数据写回原始文件

## ✨ 核心特性

### 🎯 **高保真解析**
- **多格式支持**: Excel (.xlsx, .xlsm, .xls, .xlsb)、CSV等多种格式
- **完整样式保留**: 字体、颜色、边框、对齐方式、背景色
- **结构完整性**: 合并单元格、公式和数据类型保持不变
- **大文件处理**: 自动流式处理以优化性能

### 🤖 **AI优化设计**
- **LLM友好的JSON**: 清洁、结构化的数据格式，完美适配AI处理
- **智能摘要**: 大型表格的自动数据摘要
- **灵活范围选择**: 解析特定单元格、范围或整个工作表
- **多工作表支持**: 处理包含多个工作表的复杂工作簿

### 🔧 **生产就绪**
- **MCP协议兼容**: 完全兼容Claude Desktop和其他MCP客户端
- **错误处理**: 强大的错误报告和恢复机制
- **性能优化**: 高效的内存使用和处理
- **备份支持**: 修改前自动创建文件备份

## 🛠️ 安装

### 前置要求
- Python 3.8 或更高版本
- [uv](https://docs.astral.sh/uv/) (推荐的包管理器)
- Claude Desktop 或其他兼容MCP的客户端，比如Cherry stdio

### 从源码安装
```bash
git clone https://github.com/yuqie6/MCP-Sheet-Parser.git
cd MCP-Sheet-Parser
uv sync
```

## 🚀 快速开始

### 1. 配置 Claude Desktop

将此服务器添加到您的 Claude Desktop 配置文件中：

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "sheet-parser": {
      "command": "uv",
      "args": [
        "--directory",
        "path/MCP-Sheet-Parser-cot",
        "run",
        "main.py"
      ]
    }
  }
}
```

> **注意**: 请将 `path/MCP-Sheet-Parser-cot` 替换为您的实际项目目录路径。

### 2. 其他配置选项

#### 直接使用 Python (如果没有安装 uv):
```json
{
  "mcpServers": {
    "sheet-parser": {
      "command": "python",
      "args": ["/path/to/MCP-Sheet-Parser/main.py"],
      "cwd": "/path/to/MCP-Sheet-Parser"
    }
  }
}
```

#### 使用虚拟环境:
```json
{
  "mcpServers": {
    "sheet-parser": {
      "command": "/path/to/MCP-Sheet-Parser/.venv/bin/python",
      "args": ["main.py"],
      "cwd": "/path/to/MCP-Sheet-Parser"
    }
  }
}
```

### 3. 在 Claude Desktop 中开始使用

配置完成后，重启 Claude Desktop，您就可以让 Claude 处理表格了：

> "请解析 `/path/to/sales.xlsx` 中的销售数据并显示摘要"

> "将预算表格转换为HTML，并突出显示超过10,000元的单元格"

> "更新库存表格，将所有数量增加10%"

## 🔧 可用工具

### `parse_sheet`
将表格文件解析为AI友好的JSON格式。

**参数:**
- `file_path` (必需): 表格文件的路径
- `sheet_name` (可选): 要解析的特定工作表
- `range_string` (可选): 单元格范围，如 "A1:D10"

**示例:**
```json
{
  "file_path": "/path/to/data.xlsx",
  "sheet_name": "销售数据",
  "range_string": "A1:E100"
}
```

### `convert_to_html`
将表格文件转换为保留样式的高保真HTML。

**参数:**
- `file_path` (必需): 表格文件的路径
- `output_path` (可选): HTML文件的保存位置
- `page_size` (可选): 大文件的每页行数
- `page_number` (可选): 要生成的特定页面

**示例:**
```json
{
  "file_path": "/path/to/report.xlsx",
  "output_path": "/path/to/report.html",
  "page_size": 50
}
```

### `apply_changes`
将修改后的数据写回原始表格文件。

**参数:**
- `file_path` (必需): 目标文件的路径
- `table_model_json` (必需): 从 `parse_sheet` 获取的修改后数据
- `create_backup` (可选): 写入前创建备份 (默认: true)

**示例:**
```json
{
  "file_path": "/path/to/data.xlsx",
  "table_model_json": { /* 修改后的数据 */ },
  "create_backup": true
}
```

## 🏗️ 架构

```
Claude Desktop (MCP 客户端)
           ↓
    JSON-RPC over stdin/stdout
           ↓
    MCP 表格解析服务器
           ↓
      核心服务层
           ↓
    格式特定解析器
    (XLSX, XLS, CSV, XLSB, XLSM)
```

## 🧪 开发

### 设置开发环境
```bash
git clone https://github.com/yuqie6/MCP-Sheet-Parser.git
cd MCP-Sheet-Parser
uv sync --dev
```

### 运行测试
```bash
# 使用 uv (推荐)
uv run pytest tests/ -v

# 或使用传统方法
python -m pytest tests/ -v
```

### 本地运行服务器
```bash
# 使用 uv
uv run main.py

# 或直接使用 Python
python main.py
```

### 代码结构
```
src/
├── mcp_server/          # MCP 协议实现
├── core_service.py      # 业务逻辑层
├── parsers/            # 格式特定解析器
├── converters/         # HTML 转换
├── models/             # 数据模型和工具定义
└── utils/              # 工具函数
```

## 🤝 贡献

1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/amazing-feature`)
3. 提交您的更改 (`git commit -m 'Add amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 详情请参阅 [LICENSE](LICENSE) 文件。

## 🙏 致谢

- 为 [MCP](https://modelcontextprotocol.io/) 而构建
- 专为与 Claude Desktop 无缝协作而设计
- 受到腾讯犀牛鸟计划，改善 llm 表格集成需求的启发
