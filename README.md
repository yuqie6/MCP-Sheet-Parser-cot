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
- **上下文友好**: 默认返回概览数据，避免LLM上下文爆炸
- **智能分层**: 概览→采样→完整数据，按需加载
- **参数化控制**: LLM可自主决定数据详细程度
- **使用指导**: 自动提供下一步操作建议
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

### 2. 配置cherry stdio

命令填写 **uv**

参数填写：
> --directory
> 
> path\MCP-Sheet-Parser-cot
> 
> run
> 
> main.py
> 
路径请填写绝对路径

### 3. 在 Claude Desktop 中开始使用

配置完成后，重启 Claude Desktop，您就可以让 Claude 处理表格了：

> "请解析 `/path/to/sales.xlsx` 中的销售数据并显示摘要"

> "将path/name.xlsx转换为HTML（绝对路径）"

> "更新库存表格，将所有数量增加10%"

## 🔧 可用工具

### `parse_sheet`
将表格文件解析为AI友好的JSON格式。默认返回概览信息，可按需获取详细数据。

**参数:**
- `file_path` (必需): 表格文件的绝对路径，支持 .csv, .xlsx, .xls, .xlsb, .xlsm 格式
- `sheet_name` (可选): 要解析的工作表名称，留空则使用第一个工作表
- `range_string` (可选): 单元格范围，如 "A1:D10"，指定范围时返回该范围的完整数据
- `include_full_data` (可选，默认false): 是否返回完整数据，false时只返回概览和预览
- `include_styles` (可选，默认false): 是否包含样式信息（字体、颜色、边框等）
- `preview_rows` (可选，默认5): 预览行数，当include_full_data为false时生效
- `max_rows` (可选): 最大返回行数，用于限制大文件的数据量

**使用示例:**

```json
// 基础概览（推荐首次使用）
{
  "file_path": "/path/to/data.xlsx"
}

// 获取特定工作表的完整数据
{
  "file_path": "/path/to/data.xlsx",
  "sheet_name": "销售数据",
  "include_full_data": true,
  "max_rows": 100
}

// 获取指定范围的数据（包含样式）
{
  "file_path": "/path/to/data.xlsx",
  "range_string": "A1:E50",
  "include_styles": true
}
```

### `convert_to_html`
将表格文件转换为保留样式的高保真HTML，支持多工作表和分页功能。

**参数:**
- `file_path` (必需): 源表格文件的绝对路径，支持 .csv, .xlsx, .xls, .xlsb, .xlsm 格式
- `output_path` (可选): 输出HTML文件的路径，留空则在源文件目录生成同名.html文件
- `sheet_name` (可选): 要转换的单个工作表名称，留空则转换所有工作表
- `page_size` (可选，默认100): 分页时每页显示的行数，用于控制大型文件的单页大小
- `page_number` (可选，默认1): 要查看的页码，从1开始，用于浏览大型文件的特定页面
- `header_rows` (可选，默认1): 将文件顶部的指定行数视为表头

**示例:**
```json
// 转换整个文件
{
  "file_path": "/path/to/report.xlsx",
  "output_path": "/path/to/report.html"
}

// 转换特定工作表并分页
{
  "file_path": "/path/to/large_data.xlsx",
  "sheet_name": "数据表",
  "page_size": 50,
  "page_number": 1
}
```

### `apply_changes`
将修改后的数据写回原始表格文件，完成数据编辑闭环。

**参数:**
- `file_path` (必需): 需要写回数据的目标文件的绝对路径
- `table_model_json` (必需): 从 `parse_sheet` 工具获取并修改后的 TableModel JSON 数据
  - 必须包含: `sheet_name` (字符串), `headers` (字符串数组), `rows` (二维数组)
- `create_backup` (可选，默认true): 是否在写入前创建原始文件的备份，防止意外覆盖

**示例:**
```json
{
  "file_path": "/path/to/data.xlsx",
  "table_model_json": {
    "sheet_name": "销售数据",
    "headers": ["日期", "产品", "销量", "金额"],
    "rows": [
      ["2024-01-01", "产品A", 100, 5000],
      ["2024-01-02", "产品B", 150, 7500]
    ]
  },
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

### 本地调试和测试

#### 使用MCP Inspector调试
MCP Inspector是官方提供的可视化调试工具（但是我不会用，给不了建议，贴一份官方的代码）
```
npx @modelcontextprotocol/inspector \
  uv \
  --directory path/to/server \
  run \
  package-name \
  args...
```

#### 使用客户端调试

直接要求llm使用工具并查看返回内容


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
