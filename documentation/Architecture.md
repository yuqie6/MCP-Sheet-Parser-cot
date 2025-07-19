# 项目架构文档

## 1. 概述

本项目是一个基于模型上下文协议（MCP）的电子表格处理服务器。其核心设计思想是将复杂的表格文件解析能力封装成一个独立的、可通过标准化接口（MCP）调用的服务。这种设计将 AI 代理的核心逻辑与具体的文件处理技术解耦，使得代理可以专注于任务本身，而无需关心底层文件格式的实现细节。

系统采用分层架构，确保了代码的模块化、可扩展性和可维护性。

## 2. 分层架构

系统架构主要分为以下几个层次：

```
+-----------------------------+
|   MCP 客户端 (AI 代理)      |
+-----------------------------+
             | (stdio)
             v
+-----------------------------+
|   MCP 服务器层 (mcp_server) |
+-----------------------------+
             |
             v
+-----------------------------+
|   核心服务层 (core_service) |
+-----------------------------+
             |
             v
+-----------------------------+
|   数据处理层 (parsers,      |
|   converters)             |
+-----------------------------+
             |
             v
+-----------------------------+
|   数据模型层 (models)       |
+-----------------------------+
```

1.  **MCP 服务器层 (`src/mcp_server`)**: 这是系统的入口。它负责通过标准输入输出（stdio）与 MCP 客户端进行通信，遵循 JSON-RPC 规范。该层接收请求，将其分发给相应的工具处理，并将处理结果格式化后返回给客户端。

2.  **核心服务层 (`src/core_service.py`)**: 这是业务逻辑的核心。它协调解析器、转换器和数据模型来完成具体任务。例如，`parse_sheet` 的逻辑、HTML 转换的流程以及数据写回的操作都在这一层实现。它还集成了缓存和流式读取等优化策略。

3.  **数据处理层 (`src/parsers`, `src/converters`)**: 这一层包含具体的实现细节。
    *   **解析器 (`parsers`)**: 使用工厂模式（`ParserFactory`）根据文件扩展名选择合适的解析器（如 `XLSXParser`, `CSVParser` 等），将不同的文件格式抽象为统一的内部数据结构。
    *   **转换器 (`converters`)**: 负责将内部数据结构转换为其他格式，主要是 HTML。它处理样式、合并单元格和分页等复杂的转换逻辑。

4.  **数据模型层 (`src/models`)**: 定义了系统内部使用的核心数据结构，主要是 `TableModel` (`table_model.py`)。这个标准化的内部模型（`Sheet`, `Row`, `Cell`）是各层之间传递数据的通用语言。此外，该层还定义了暴露给 MCP 客户端的工具（`tools.py`）。

## 3. 核心组件详解

-   **`main.py`**: 项目的启动入口，负责初始化并运行 MCP 服务器。

-   **`mcp_server/server.py`**:
    -   使用官方 `mcp.server` 库创建服务器实例。
    -   通过 `stdio_server` 建立与客户端的通信通道。
    -   调用 `register_tools` 注册所有可用的工具。

-   **`models/tools.py`**:
    -   定义了 `parse_sheet`, `convert_to_html`, `apply_changes` 三个核心工具的 MCP 接口，包括它们的名称、描述和输入/输出模式（Schema）。
    -   作为工具调用的分发器，将请求路由到 `CoreService` 中对应的处理方法。

-   **`core_service.py`**:
    -   `CoreService` 类是所有业务逻辑的集合点。
    -   `parse_sheet_optimized`: 实现了为 AI 代理优化的解析逻辑，能够根据参数返回概览、完整数据或特定范围的数据。
    -   `convert_to_html`: 调用 `HTMLConverter` 或 `PaginatedHTMLConverter` 执行转换。
    -   `apply_changes`: 实现将修改后的数据写回到不同格式文件的逻辑。
    -   集成了缓存（`src/cache`）和流式读取（`src/streaming`）等高级功能。

-   **`parsers/factory.py`**:
    -   `ParserFactory` 类根据文件后缀名（`.xlsx`, `.csv` 等）动态选择并实例化相应的解析器。这使得添加对新文件格式的支持变得简单，只需实现一个新的解析器类即可。

-   **`models/table_model.py`**:
    -   定义了 `Sheet`, `Row`, `Cell`, `Style` 等 `pydantic` 模型。这些模型构成了系统内部处理和传递数据的标准格式，确保了数据的一致性。

## 4. 数据流示例 (`parse_sheet`)

1.  AI 代理（客户端）向服务器发送一个 `parse_sheet` 的 JSON-RPC 请求。
2.  `mcp_server/server.py` 接收到请求，并将其传递给 `models/tools.py` 中的 `handle_call_tool` 函数。
3.  `handle_call_tool` 识别出工具名称为 `parse_sheet`，调用 `_handle_parse_sheet` 辅助函数。
4.  `_handle_parse_sheet` 调用 `core_service.parse_sheet_optimized` 方法，并传入所有参数。
5.  `CoreService` 首先通过 `ParserFactory` 获取适用于目标文件类型的解析器。
6.  解析器读取文件，将其内容转换为内部标准的 `Sheet` 对象。
7.  `CoreService` 根据 `include_full_data` 等参数，对 `Sheet` 对象进行处理，提取概览信息或完整数据。
8.  处理结果被格式化为 JSON，并逐层返回，最终通过 stdio 发送回 AI 代理。

## 5. 目录结构

```
src/
├── cache/            # 缓存实现 (LRU, Disk)
├── config/           # 配置文件
├── converters/       # 将内部数据模型转换为其他格式 (如 HTML)
├── mcp_server/       # MCP 协议实现和服务器入口
├── models/           # 数据模型 (TableModel) 和工具定义 (Tools)
├── parsers/          # 针对不同文件格式的解析器
├── streaming/        # 大文件流式读取的实现
└── utils/            # 通用工具函数 (颜色处理, 样式解析等)