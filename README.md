# MCP-Sheet-Parser: AI 原生表格处理服务


<p align="center">
  <strong>一个专为 AI 助力量身打造的高性能、高保真 MCP 服务，赋予 AI 精确读写、修改、转换复杂表格文件的能力。</strong>
</p>

<p align="center">
  <a href="#核心功能">核心功能</a> •
  <a href="#技术亮点">技术亮点</a> •
  <a href="#快速开始">快速开始</a> •
  <a href="#工具参考">工具参考</a>
</p>

---

**MCP-Sheet-Parser** 是一个完整的表格处理解决方案，旨在解决 AI 助手在处理电子表格时的核心痛点。它通过提供一套标准化的工具接口，让 AI 能够像人类专家一样，无缝地处理 `CSV`, `XLSX`, `XLS` 等多种主流表格格式，并实现了对复杂样式的精确解析与回写。

MCP-Sheet-Parser-Cot 是一个功能强大的高保真电子表格解析器与转换器，旨在作为一个标准的模型上下文协议（MCP）服务器工具运行。它赋予了AI大语言模型（LLM）以编程方式读取、理解甚至修改多种电子表格文件（包括 XLSX, XLS, CSV 等）的能力。

**核心特性包括：**
*   **广泛的格式支持**：覆盖主流的Excel和CSV文件格式。
*   **高保真样式还原**：精确映射字体、颜色、背景、对齐、边框等样式到HTML。
*   **复杂结构处理**：正确处理合并单元格、超链接和单元格注释。
*   **为AI设计的闭环工具链**：提供 `convert_to_html` (查看)、`parse_sheet` (理解) 和 `apply_changes` (修改) 的完整交互闭环。
*   **大文件优化**：内置分页和范围选择功能，高效处理大型数据集。

## 核心功能

- **多格式无缝处理**: 统一的接口，轻松处理 `CSV`, `XLSX`, `XLS`, `XLSB`, `XLSM` 等多种文件格式。
- **高保真样式解析**: 精确解析**字体、颜色、边框、对齐、数字格式、超链接、注释**等复杂样式。
- **端到端数据闭环**: 提供从文件解析 (`parse_sheet`)、数据修改到变更写回 (`apply_changes`) 的完整工作流，实现真正的程序化编辑。
- **可视化转换**: 将表格文件转换为高保真的 HTML 进行预览，支持分页。

## 技术亮点

- **为超大文件优化**:
  - **流式读取 (Streaming)**: 对 `CSV` 和 `XLSX` 等格式支持流式处理，通过 `LazySheet` 模式，即使是百万行级别的超大文件也能以极低的内存占用进行解析。
  - **智能缓存 (Caching)**: 内置二级缓存系统（内存 L1 + 磁盘 L2），并采用基于文件内容、大小和修改时间的智能哈希算法作为缓存键，确保在文件未变更时实现毫秒级响应。
- **高性能HTML生成**:
  - **智能CSS生成**: 通过动态生成CSS类而非内联样式，极大地减小了HTML文件体积，并提升了浏览器的渲染性能。
  - **分页渲染**: 对大表格的HTML转换内置分页功能，实现了快速加载和流畅预览。
- **LLM 友好型数据结构**: 将表格解析为专为大语言模型优化的 `TableModel` JSON 格式，结构清晰，便于 AI 理解、分析和修改。
- **清晰的模块化架构**:
  - **服务层 (`mcp_server/`)**: 统一的 MCP 工具入口。
  - **核心服务层 (`core_service.py`)**: 协调下层模块完成任务。
  - **解析器层 (`parsers/`)**: 采用工厂模式，每种格式由专门的解析器处理，易于扩展。
  - **数据模型层 (`models/`)**: 定义了 `Sheet`, `LazySheet` 等标准化的数据结构。
  - **转换器层 (`converters/`)**: 负责将 `TableModel` 转换为 HTML 等其他格式。

## 快速开始

### 环境要求
- Python 3.8+
- [uv](https://github.com/astral-sh/uv) (推荐的包管理器)

### 安装与运行
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
将此服务集成到 AI 助手（如 Claude Desktop）的 `mcpServers` 配置中：
```json
{
  "sheetParser": {
    "command": "uv",
    "args": ["run", "python", "-m", "src.mcp_server.server"],
    "cwd": "/path/to/your/MCP-Sheet-Parser"
  }
}
```

## 工具参考 (MCP Tools)

### 1. `parse_sheet`
将表格文件智能解析为结构化的 `TableModel` JSON 对象。
- **参数**:
  - `file_path` (string, **必需**): 目标表格文件的绝对路径。
  - `sheet_name` (string, *可选*): 要解析的工作表名称。
  - `range_string` (string, *可选*): 要解析的单元格范围 (e.g., `'A1:D10'`)。

### 2. `apply_changes`
将修改后的 `TableModel` JSON 数据写回到原始文件中。
- **参数**:
  - `file_path` (string, **必需**): 要写回的目标文件的绝对路径。
  - `table_model_json` (object, **必需**): 从 `parse_sheet` 获取并修改后的 `TableModel` JSON 对象。
  - `create_backup` (boolean, *可选*): 是否创建备份文件，默认为 `true`。

### 3. `convert_to_html`
将表格文件转换为高保真、可交互的 HTML 格式。
- **参数**:
  - `file_path` (string, **必需**): 源表格文件的绝对路径。
  - `output_path` (string, *可选*): 输出 HTML 文件的路径。
  - `page_size` (integer, *可选*): 分页时每页的行数，默认为 `100`。
  - `page_number` (integer, *可选*): 要查看的页码，从 `1` 开始。

---

*要了解项目当前已知的设计优化点和待办事项，请参阅 `docs/IMPROVEMENTS.md`。*
