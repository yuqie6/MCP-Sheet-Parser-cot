# MCP-Sheet-Parser: AI 原生表格处理服务

<p align="center">
  <img src="https://raw.githubusercontent.com/user-attachments/assets/dd33663a-8f9f-4a94-b1f4-30740a33a39e" alt="MCP-Sheet-Parser Logo" width="150">
</p>

<p align="center">
  <strong>一个专为 AI 助力量身打造的高性能、高保真 MCP 服务，赋予 AI 精确读写、修改、转换复杂表格文件的能力。</strong>
</p>

<p align="center">
  <a href="#核心优势">核心优势</a> •
  <a href="#架构设计">架构设计</a> •
  <a href="#安装与配置">安装与配置</a> •
  <a href="#工具参考">工具参考</a> •
  <a href="#开发与测试">开发与测试</a>
</p>

---

**MCP-Sheet-Parser** 不仅仅是一个文件转换器，它是一个完整的表格处理解决方案，旨在解决 AI 助手在处理电子表格时的核心痛点。它通过提供一套标准化的工具接口，让 AI 能够像人类专家一样，无缝地处理 `CSV`, `XLSX`, `XLS` 等多种主流表格格式，并实现了对复杂样式的精确解析与回写。

## 核心优势

- 🚀 **多格式无缝处理**: 统一的接口，轻松处理 `CSV`, `XLSX`, `XLS`, `XLSB`, `XLSM` 等多种文件格式。
- ✨ **复杂样式高保真**: 业界领先的样式处理引擎，能精确解析并写回 **字体、颜色、边框、对齐、数字格式、超链接、注释** 等复杂样式，保持表格原有外观。
- ⚡ **为超大文件优化**:
  - **流式读取 (Streaming)**: 对 `CSV`, `XLSX`, `XLSM` 格式支持流式处理，通过 `LazySheet` 和 `RowProvider` 模式，即使是百万行级别的超大文件也能以极低的内存占用进行解析。
  - **分页预览 (Pagination)**: `convert_to_html` 工具内置分页功能，将大表格自动切片，实现快速加载和流畅预览。
- 🤖 **LLM 友好型数据结构**: 将表格解析为专为大语言模型优化的 `TableModel` JSON 格式，结构清晰，便于 AI 理解、分析和修改。
- 🔄 **端到端数据闭环**: 提供从文件解析 (`parse_sheet`)、数据修改到变更写回 (`apply_changes`) 的完整工作流，实现真正的程序化、无损编辑。

## 架构设计

项目采用清晰的、可扩展的分层架构，确保了代码的模块化和高内聚性。

1.  **MCP 服务层 (`mcp_server/`)**: 作为项目入口，负责监听和响应 AI 助手的 MCP 工具调用请求。`tools.py` 在此层定义了三个核心工具。
2.  **核心服务层 (`core_service.py`)**: 系统的“大脑”，扮演着外观（Facade）和协调者的角色。它接收来自服务层的请求，并智能地调度下层解析器、转换器来完成具体任务。
3.  **解析器层 (`parsers/`)**:
    - **解析器工厂 (`factory.py`)**: 根据文件扩展名动态选择并实例化最合适的解析器。
    - **具体解析器 (`xlsx_parser.py`, etc.)**: 每个解析器专职处理一种文件格式，负责将文件流转换为标准化的 `TableModel`。`XlsxParser` 等高级解析器内置了复杂的样式提取逻辑。
4.  **数据模型层 (`models/`)**:
    - **`table_model.py`**: 定义了整个系统的核心数据结构，如 `Sheet`, `Row`, `Cell` 和 `Style`。这个标准化的模型是所有操作的基础，确保了数据在各层之间的一致性。
    - **`LazySheet` 和 `LazyRowProvider`**: 在此定义，为流式处理提供了契约。
5.  **转换器层 (`converters/`)**: 负责将内存中的 `TableModel` 对象转换为其他格式，如 `html_converter.py` 将其转换为 HTML 进行预览。

这种设计使得添加对新文件格式的支持变得简单——只需实现一个新的解析器并注册到工厂中即可。

## 安装与配置

### 环境要求
- Python 3.8+
- [uv](https://github.com/astral-sh/uv) (推荐的包管理器)

### 安装步骤

```bash
# 1. 克隆项目
git clone https://github.com/your-repo/MCP-Sheet-Parser.git
cd MCP-Sheet-Parser

# 2. 安装依赖
uv install

# 3. 运行测试以验证安装
uv run python -m pytest tests/
```

### AI 助手配置

要将此服务集成到 AI 助手中 (例如 Claude Desktop)，请在助手的配置文件中添加 MCP 服务器配置：

```json
{
  "mcpServers": {
    "sheetParser": {
      "command": "uv",
      "args": ["run", "python", "-m", "src.mcp_server.server"],
      "cwd": "/path/to/your/MCP-Sheet-Parser"
    }
  }
}
```

## 工具参考 (MCP Tools)

### 1. `convert_to_html`

将表格文件转换为高保真、可交互的 HTML 格式。

- **功能**:
  - 精确还原单元格样式，包括字体、颜色、背景、边框等。
  - 自动为大型文件启用分页，提升加载速度和用户体验。
- **参数**:
  - `file_path` (string, **必需**): 源表格文件的绝对路径。
  - `output_path` (string, *可选*): 输出 HTML 文件的路径。如果留空，将在源文件同目录下生成同名 `.html` 文件。
  - `page_size` (integer, *可选*): 分页时每页的行数，默认为 `100`。
  - `page_number` (integer, *可选*): 要查看的页码，从 `1` 开始，默认为 `1`。

### 2. `parse_sheet`

将表格文件智能解析为结构化的 `TableModel` JSON 对象。

- **功能**:
  - 提取数据及 **完整的样式信息**。
  - 支持流式处理，对于大文件可返回摘要或通过 `range_string` 进行局部加载。
  - 可选择特定的工作表或单元格范围进行解析。
- **参数**:
  - `file_path` (string, **必需**): 目标表格文件的绝对路径。
  - `sheet_name` (string, *可选*): 要解析的工作表名称。如果留空，则使用默认的活动工作表。
  - `range_string` (string, *可选*): 要解析的单元格范围 (e.g., `'A1:D10'`)。在处理大文件时强烈推荐使用此参数。

### 3. `apply_changes`

将修改后的 `TableModel` JSON 数据写回到原始文件中。

- **功能**:
  - 将 `TableModel` 中的数据和样式变更精确应用到目标文件。
  - **保留原始文件的格式和未被修改部分的样式与内容**。
  - 默认自动创建文件备份，确保数据安全。
- **参数**:
  - `file_path` (string, **必需**): 要写回的目标文件的绝对路径。
  - `table_model_json` (object, **必需**): 从 `parse_sheet` 获取并修改后的 `TableModel` JSON 对象。
  - `create_backup` (boolean, *可选*): 是否创建备份文件，默认为 `true`。

## 开发与测试

### 运行测试
项目使用 `pytest` 进行全面的单元测试和集成测试。
```bash
# 运行所有测试
uv run python -m pytest tests/ -v

# 运行特定的端到端测试
uv run python -m pytest tests/test_end_to_end.py -v

# 查看测试覆盖率报告
uv run python -m pytest tests/ --cov=src --cov-report=term-missing
```

### 贡献指南
我们欢迎社区的任何贡献！
1.  **Fork** 本项目。
2.  创建您的特性分支 (`git checkout -b feature/AmazingFeature`)。
3.  提交您的更改 (`git commit -m 'Add some AmazingFeature'`)。
4.  确保所有测试都通过。
5.  将您的分支推送到远程 (`git push origin feature/AmazingFeature`)。
6.  **提交一个 Pull Request**，并清晰地描述您的工作。

---

MIT License © 2024 [Your Name]
