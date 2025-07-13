# MCP-Sheet-Parser-Cot

![PyPI - Version](https://img.shields.io/pypi/v/MCP-Sheet-Parser-Cot)
![PyPI - License](https://img.shields.io/pypi/l/MCP-Sheet-Parser-Cot)
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/MCP-Sheet-Parser-Cot)

一个为AI智能体设计的高保真电子表格解析与HTML转换器 (MCP工具)。

**MCP-Sheet-Parser-Cot** 是一个功能强大的高保真电子表格解析器与转换器，旨在作为一个标准的模型上下文协议（MCP）服务器工具运行。它赋予了AI大语言模型（LLM）以编程方式读取、理解甚至修改多种电子表格文件（包括 `.xlsx`, `.xls`, `.csv` 等）的能力。

## 核心特性

- **高保真解析**: 精确解析 `.xlsx`, `.xls`, `.csv` 等多种格式的电子表格，保留单元格的原始数据类型、样式（字体、颜色、边框、对齐方式）和合并单元格结构。
- **结构化数据模型**: 将复杂的电子表格数据转换为一个标准化的、易于AI理解的 `TableModel` JSON结构。
- **HTML 转换**: 将 `TableModel` 渲染为高度还原的HTML，方便在Web界面中进行预览和展示。
- **数据回写**: 支持将AI修改后的 `TableModel` 数据安全地回写到原始电子表格文件中。
- **MCP兼容**: 作为标准的MCP (Model Context Protocol) 服务器运行，使AI智能体能通过工具调用来操作电子表格。

## 架构概览

- **分层设计**: 采用 `MCP Server` -> `Core Service` -> `Parsers` 的分层架构。
  - **MCP Server (`src/mcp_server`)**: 对外暴露工具接口。
  - **Core Service (`src/core_service.py`)**: 核心业务逻辑，调度不同格式的解析器。
  - **Parsers (`src/parsers`)**: 每种文件格式对应一个独立的解析器，易于扩展。
- **核心数据模型**: 所有操作都围绕 `TableModel` (`src/models/table_model.py`) 进行，实现数据在不同模块间的解耦和标准化。

## 安装

```bash
pip install MCP-Sheet-Parser-Cot
```
*注意：包尚未发布，此命令将在发布后可用。*

## 使用方法

### 作为MCP服务器运行
安装包后，通过以下命令启动服务器：
```bash
mcp-sheet-parser
# 或者指定端口
mcp-sheet-parser --port 8080
```

### 工具调用示例 (Tool Call Examples)

以下示例展示了如何通过 `curl` 与运行中的MCP服务器进行交互。

#### 1. `parse_sheet`: 解析电子表格为 TableModel

```bash
# Request
curl -X POST http://127.0.0.1:8000/parse_sheet -H "Content-Type: application/json" -d '{
  "file_path": "tests/test_data/complex_styles.xlsx",
  "sheet_name": "Sheet1"
}'
```
**响应**: 返回一个包含工作表结构、样式和数据的 `TableModel` JSON 对象。

#### 2. `convert_to_html`: 将 TableModel 转换为HTML
为了演示，我们将上一步 `parse_sheet` 的输出作为 `convert_to_html` 的输入。

```bash
# Request
# 注意: 'table_model' 的内容是上一步 'parse_sheet' 响应的完整JSON
curl -X POST http://127.0.0.1:8000/convert_to_html -H "Content-Type: application/json" -d '{
  "table_model": {
    "sheet_name": "Sheet1",
    "headers": [
        {"value": "ID", "style": {"font": {"bold": true}}},
        {"value": "Name", "style": {"font": {"bold": true}}},
        {"value": "Value", "style": {"font": {"bold": true}}}
    ],
    "rows": [
        [
            {"value": 1, "style": {}},
            {"value": "Item A", "style": {}},
            {"value": 100, "style": {"fill": {"fgColor": "FFFF00"}}}
        ],
        [
            {"value": 2, "style": {}},
            {"value": "Item B", "style": {}},
            {"value": 200, "style": {"font": {"color": "FF0000"}}}
        ]
    ],
    "styles": [],
    "merged_cells": []
  }
}'
```
**响应**: 返回一个HTML字符串，用于在浏览器中显示表格。

#### 3. `apply_changes`: 将修改后的 TableModel 回写到文件

```bash
# Request
# 假设我们修改了 Item A 的值为 150
curl -X POST http://127.0.0.1:8000/apply_changes -H "Content-Type: application/json" -d '{
  "file_path": "tests/test_data/complex_styles.xlsx",
  "table_model": {
    "sheet_name": "Sheet1",
    "headers": [
        {"value": "ID", "style": {"font": {"bold": true}}},
        {"value": "Name", "style": {"font": {"bold": true}}},
        {"value": "Value", "style": {"font": {"bold": true}}}
    ],
    "rows": [
        [
            {"value": 1, "style": {}},
            {"value": "Item A", "style": {}},
            {"value": 150, "style": {"fill": {"fgColor": "FFFF00"}}}
        ],
        [
            {"value": 2, "style": {}},
            {"value": "Item B", "style": {}},
            {"value": 200, "style": {"font": {"color": "FF0000"}}}
        ]
    ],
    "styles": [],
    "merged_cells": []
  }
}'
```
**响应**: `{"success": true, "message": "Changes applied successfully."}`

## AI 助手配置
将此服务集成到 AI 助手（如 Claude Desktop）的 `mcpServers` 配置中：
```json
{
  "sheetParser": {
    "command": "mcp-sheet-parser"
  }
}
```

## 许可证
本项目采用 [MIT](LICENSE) 许可证。
