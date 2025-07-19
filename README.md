# MCP Sheet Parser

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io/)

MCP Sheet Parser 是一个基于模型上下文协议（Model Context Protocol, MCP）构建的服务器。它为 AI 代理提供解析、转换和修改电子表格文件的能力。

该服务器通过标准输入输出（stdio）与兼容 MCP 的客户端（如 Claude Desktop）通信，允许 AI 代理以工具调用的方式处理包括 XLSX、XLSM、XLS、XLSB 和 CSV 在内的多种表格格式。

## 工作原理

当 AI 代理需要处理一个表格文件时，它会通过 MCP 向本服务器发送一个 JSON-RPC 请求。服务器接收请求，调用内部相应的处理函数，并将结构化的 JSON 数据返回给代理。这个过程使得 AI 代理能够理解和操作传统上难以访问的表格数据。

其基本架构如下：
1.  **MCP 客户端**: AI 代理的运行环境，例如 Claude Desktop 或其他 IDE。
2.  **通信协议**: 客户端与服务器之间通过标准输入输出（stdio）进行 JSON-RPC 通信。
3.  **MCP 服务器**: 本项目，一个独立的 Python 进程，负责监听和响应来自客户端的请求。
4.  **核心服务**: 服务器内部的业务逻辑，调用特定的解析器或转换器来完成任务。

## 核心功能

服务器提供三个核心工具来完成一个完整的数据处理闭环：

1.  **`parse_sheet`**: 解析电子表格文件。此工具将文件内容转换为结构化的 JSON 对象，该对象为 AI 代理的上下文进行了优化。默认情况下，它只返回文件的概览信息（如尺寸、列名和数据预览），以避免消耗过多的令牌。代理可以根据需要请求获取完整数据或样式信息。

2.  **`convert_to_html`**: 将电子表格转换为 HTML 文件。此功能可以保留原始文件中的大部分样式，包括字体、颜色、边框和合并单元格，使得数据可以在浏览器中进行可视化查阅。

3.  **`apply_changes`**: 将 AI 代理修改后的 JSON 数据写回到原始电子表格文件中。此工具接收从 `parse_sheet` 获取并由代理处理过的数据，完成数据的修改和保存。

## 安装与配置

### 前置要求
- Python 3.8 或更高版本
- [uv](https://docs.astral.sh/uv/) (推荐的包管理器)

### 安装步骤
```bash
git clone https://github.com/yuqie6/MCP-Sheet-Parser.git
cd MCP-Sheet-Parser
uv sync
```

### 客户端配置
要将此服务器与兼容 MCP 的客户端一同使用，需要在客户端的配置文件中进行设置。

以 Claude Desktop 为例，配置文件路径如下：
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%/Claude/claude_desktop_config.json`

在配置文件中添加以下 `mcpServers` 条目：
```json
{
  "mcpServers": {
    "sheet-parser": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/MCP-Sheet-Parser-cot",
        "run",
        "main.py"
      ]
    }
  }
}
```
**注意**: 请将 `/path/to/MCP-Sheet-Parser-cot` 替换为本项目的绝对路径。

## 使用指南

配置完成后，重启客户端即可开始使用。您可以向 AI 代理发出自然语言指令来处理表格文件。

- **解析文件**: "请解析 `/path/to/sales.xlsx` 文件并提供一份摘要。"
- **转换文件**: "将 `/path/to/data.csv` 转换为 HTML 文件。"
- **修改数据**: "读取库存表 `/path/to/inventory.xlsx`，然后将所有'笔记本电脑'的数量增加10，最后保存修改。"

## 工具定义

### `parse_sheet`
解析一个表格文件，返回其结构化的 JSON 表示。

- **`file_path`** (字符串, 必需): 表格文件的绝对路径。
- **`sheet_name`** (字符串, 可选): 需要解析的特定工作表名称。如果留空，则解析第一个工作表。
- **`range_string`** (字符串, 可选): 指定单元格范围，例如 "A1:D10"。
- **`include_full_data`** (布尔值, 可选, 默认 `false`): 是否返回所有行的数据。
- **`include_styles`** (布尔值, 可选, 默认 `false`): 是否在返回的数据中包含样式信息。
- **`preview_rows`** (整数, 可选, 默认 `5`): 在概览模式下，返回的数据预览行数。
- **`max_rows`** (整数, 可选): 限制返回的最大行数，用于处理大型文件。

### `convert_to_html`
将一个表格文件转换为 HTML。

- **`file_path`** (字符串, 必需): 源表格文件的绝对路径。
- **`output_path`** (字符串, 可选): 输出 HTML 文件的路径。如果留空，则在源文件相同目录下生成同名 HTML 文件。
- **`sheet_name`** (字符串, 可选): 指定要转换的单个工作表名称。如果留空，则转换所有工作表。
- **`page_size`** (整数, 可选, 默认 `100`): 分页时每页显示的行数。
- **`page_number`** (整数, 可选, 默认 `1`): 查看分页结果时的页码。
- **`header_rows`** (整数, 可选, 默认 `1`): 将文件顶部的指定行数视为固定表头。

### `apply_changes`
将修改后的数据写回表格文件。

- **`file_path`** (字符串, 必需): 目标文件的绝对路径。
- **`table_model_json`** (对象, 必需): 从 `parse_sheet` 工具获取并由 AI 代理修改后的数据对象。
- **`create_backup`** (布尔值, 可选, 默认 `true`): 是否在写入前创建原始文件的备份。

## 许可证

本项目采用 MIT 许可证。详情请参阅 [LICENSE](LICENSE) 文件。
