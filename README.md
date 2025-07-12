# MCP Sheet Parser

一个专门为 AI 助手设计的 Model Context Protocol (MCP) 服务器，让 AI 能够直接处理表格文件。

## 什么是 MCP Sheet Parser

这是一个 **MCP 服务器**，为 Claude、GPT 等 AI 助手提供表格文件处理能力。通过 MCP 协议，AI 可以：

- 直接读取和解析各种表格文件
- 将表格转换为 HTML 进行展示
- 修改表格数据并保存回原文件

**关键特点：** AI 助手可以像使用内置功能一样使用这些表格处理工具。

## MCP 工具

本服务器为 AI 助手提供三个专用工具：

### 1. convert_to_html
**功能：** 将表格文件转换为 HTML 格式，便于 AI 向用户展示表格内容
**使用场景：** 当用户要求查看表格文件时，AI 可以转换为 HTML 并展示

### 2. parse_sheet
**功能：** 将表格文件解析为结构化 JSON 数据，便于 AI 分析和处理
**使用场景：** 当 AI 需要分析表格数据、进行计算或提取信息时

### 3. apply_changes
**功能：** 将 AI 处理后的数据写回原表格文件
**使用场景：** 当用户要求修改表格数据时，AI 可以直接更新文件

### 文件格式支持

| 格式 | 读取 | HTML转换 | JSON解析 | 数据写回 | 备注 |
|------|------|----------|----------|----------|------|
| CSV | ✅ | ✅ | ✅ | ✅ | 完全支持 |
| XLSX | ✅ | ✅ | ✅ | ✅ | 基本样式支持 |
| XLSM | ✅ | ✅ | ✅ | ✅ | 基本样式支持 |
| XLS | ✅ | ✅ | ✅ | ❌ | 只读支持 |
| XLSB | ✅ | ✅ | ✅ | ❌ | 只读支持 |

## 安装和配置

### 环境要求
- Python 3.8+
- 建议使用 `uv` 包管理器

### 安装步骤
```bash
# 克隆项目
git clone <repository-url>
cd MCP-Sheet-Parser

# 安装依赖
uv install

# 验证安装
uv run python -m pytest tests/test_end_to_end.py -v
```

### 配置 AI 助手

#### 对于 Claude Desktop
在 Claude Desktop 的配置文件中添加：
```json
{
  "mcpServers": {
    "sheet-parser": {
      "command": "uv",
      "args": ["run", "python", "-m", "src.mcp_server.server"],
      "cwd": "/path/to/MCP-Sheet-Parser"
    }
  }
}
```

#### 手动启动服务器
```bash
uv run python -m src.mcp_server.server
```

## AI 助手使用示例

### 典型对话场景

**用户：** "请帮我查看这个 sales_data.xlsx 文件的内容"

**AI 助手会：**
1. 使用 `convert_to_html` 工具将文件转换为 HTML
2. 向用户展示表格内容

**用户：** "请将第二行的销售额改为 5000"

**AI 助手会：**
1. 使用 `parse_sheet` 工具读取当前数据
2. 修改指定的数据
3. 使用 `apply_changes` 工具保存修改

### MCP 工具调用示例

#### convert_to_html 工具
```json
{
  "file_path": "data.xlsx",
  "output_path": "output.html"
}
```

返回：
```json
{
  "status": "success",
  "output_path": "output.html",
  "file_size": 1234,
  "rows_converted": 10,
  "cells_converted": 40
}
```

#### 2. parse_sheet
解析表格文件为JSON格式。

**输入参数：**
```json
{
  "file_path": "data.csv",
  "range_string": "A1:D10"  // 可选
}
```

**输出示例：**
```json
{
  "sheet_name": "Sheet1",
  "headers": ["Name", "Age"],
  "rows": [
    [{"value": "Alice"}, {"value": "25"}],
    [{"value": "Bob"}, {"value": "30"}]
  ],
  "metadata": {
    "total_rows": 3,
    "total_cols": 2
  }
}
```

#### 3. apply_changes
将修改写回原文件。

**输入参数：**
```json
{
  "file_path": "data.csv",
  "table_model_json": {
    "sheet_name": "Sheet1",
    "headers": ["Name", "Age"],
    "rows": [
      [{"value": "Alice"}, {"value": "26"}]
    ]
  },
  "create_backup": true
}
```

**输出示例：**
```json
{
  "status": "success",
  "message": "数据修改已成功应用",
  "changes_applied": 1,
  "backup_created": true,
  "backup_path": "data.csv.backup"
}
```

## 限制和注意事项

### 文件大小
- 大文件（>1000行）处理可能较慢
- 建议分批处理超大文件

### 样式支持
- Excel样式支持有限，主要保留基本格式
- 复杂样式可能丢失
- 宏和公式不会保留

### 数据类型
- JSON中所有数据以字符串存储
- 写回时尝试自动类型转换
- 特殊格式（如日期）可能需要手动处理

### 错误处理
- 文件损坏或格式错误会返回错误信息
- 建议在生产环境中添加额外的验证

## 测试

### 运行测试
```bash
# 完整测试套件
uv run python -m pytest tests/ -v

# 端到端测试
uv run python -m pytest tests/test_end_to_end.py -v

# 数据写回测试
uv run python -m pytest tests/test_data_writeback.py -v

# 查看测试覆盖率
uv run python -m pytest tests/ --cov=src --cov-report=term-missing
```

### 测试状态
- 总测试数：131个
- 通过率：约95%
- 代码覆盖率：约57%

## 项目结构

```
src/
├── core_service.py          # 核心业务逻辑
├── models/                  # 数据模型定义
│   └── table_model.py
├── parsers/                 # 文件解析器
│   ├── base_parser.py
│   ├── csv_parser.py
│   ├── xlsx_parser.py
│   └── ...
├── converters/              # 格式转换器
│   └── html_converter.py
└── mcp_server/             # MCP服务器
    ├── server.py
    └── tools.py

tests/                      # 测试文件
├── test_end_to_end.py      # 端到端测试
├── test_data_writeback.py  # 数据写回测试
└── ...
```

## 开发状态

### 已完成功能
- ✅ 基本文件解析（CSV, XLSX, XLSM, XLS, XLSB）
- ✅ HTML转换
- ✅ JSON格式输出
- ✅ CSV和XLSX数据写回
- ✅ 自动备份机制
- ✅ 基本错误处理

### 待改进功能
- ⚠️ XLS/XLSB格式写回支持
- ⚠️ 复杂Excel样式处理
- ⚠️ 大文件性能优化
- ⚠️ 更完善的错误处理

### 已知问题
- 某些Excel样式可能无法完全保留
- 大文件处理性能有待优化
- 部分边界情况的错误处理需要完善

## 贡献指南

1. 提交Issue前请先检查现有Issue
2. 新功能请包含相应测试
3. 确保测试通过后再提交PR
4. 遵循现有代码风格

## 许可证

[待定]

---

**注意：** 本项目仍在开发中，建议在生产环境使用前进行充分测试。
