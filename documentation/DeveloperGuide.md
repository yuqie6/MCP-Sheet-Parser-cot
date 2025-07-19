# 开发者指南

## 1. 简介

本指南为需要对 MCP Sheet Parser 进行维护、扩展或二次开发的开发者提供技术参考。在开始之前，请确保您已熟悉《项目架构文档》中描述的系统分层和核心组件。

## 2. 环境设置

### 2.1 克隆仓库
```bash
git clone https://github.com/yuqie6/MCP-Sheet-Parser.git
cd MCP-Sheet-Parser
```

### 2.2 安装依赖
项目使用 `uv` 作为包管理器以提高依赖安装速度。
```bash
# 安装生产和开发依赖
uv sync --dev
```
此命令会读取 `pyproject.toml` 文件并创建一个虚拟环境，然后安装所有必需的库，包括 `pytest` 等开发工具。

### 2.3 运行测试
项目包含一套单元测试和集成测试，用于验证核心功能的正确性。
```bash
# 运行所有测试
uv run pytest tests/
```
在进行任何代码修改后，请务必运行完整的测试套件，确保没有引入新的问题。

## 3. 本地调试

### 3.1 直接运行服务器
您可以直接从命令行启动服务器，以便进行调试。
```bash
uv run main.py
```
服务器将启动并监听标准输入。此时，您可以从另一个终端或通过 MCP 客户端向其发送 JSON-RPC 请求。

### 3.2 使用 MCP Inspector
`@modelcontextprotocol/inspector` 是官方提供的可视化调试工具，可以帮助您查看客户端与服务器之间的通信详情。
```bash
# 启动服务器并连接到 Inspector
npx @modelcontextprotocol/inspector \
  uv \
  --directory . \
  run \
  main.py
```
**注意**: 您需要先安装 Node.js 和 npx。

## 4. 代码扩展

### 4.1 添加新的文件格式支持

要添加对一种新文件格式（例如 `.ods`）的支持，您需要完成以下步骤：

1.  **创建新的解析器**:
    在 `src/parsers/` 目录下，创建一个新的解析器文件，例如 `ods_parser.py`。
    ```python
    # src/parsers/ods_parser.py
    from .base_parser import BaseParser
    from ..models.table_model import Sheet

    class ODSParser(BaseParser):
        def parse(self, file_path: str) -> list[Sheet]:
            # 在这里实现您的 .ods 文件解析逻辑
            # ...
            # 您需要将文件内容转换为一个或多个 Sheet 对象的列表
            pass
    ```

2.  **注册新的解析器**:
    在 `src/parsers/factory.py` 文件中，将新的解析器添加到 `ParserFactory` 中。
    ```python
    # src/parsers/factory.py
    from .xlsx_parser import XLSXParser
    from .csv_parser import CSVParser
    from .ods_parser import ODSParser # 导入新的解析器

    class ParserFactory:
        def __init__(self):
            self._parsers = {
                '.xlsx': XLSXParser(),
                '.xlsm': XLSXParser(),
                '.xls': XLSXParser(), # 假设可以复用
                '.xlsb': XLSXParser(), # 假设可以复用
                '.csv': CSVParser(),
                '.ods': ODSParser(), # 注册新的解析器
            }
        # ...
    ```

3.  **（可选）实现写回逻辑**:
    如果您希望支持将数据写回到新格式的文件中，需要在 `src/core_service.py` 的 `apply_changes` 方法中添加相应的处理逻辑。

### 4.2 添加新的工具

1.  **定义工具接口**:
    在 `src/models/tools.py` 的 `handle_list_tools` 函数中，添加新工具的定义，包括名称、描述和输入模式（Input Schema）。

2.  **实现工具逻辑**:
    -   在 `src/core_service.py` 中为新工具创建一个新的方法。
    -   在 `src/models/tools.py` 的 `handle_call_tool` 函数中，添加一个新的 `elif` 分支，将工具调用路由到您在 `CoreService` 中创建的新方法。

## 5. 核心模块依赖

-   **`mcp.server`**: 用于创建和管理 MCP 服务器的基础库。
-   **`openpyxl`**: 用于处理 `.xlsx` 和 `.xlsm` 文件。
-   **`xlrd`**: 用于处理旧版的 `.xls` 文件。
-   **`pyxlsb`**: 用于处理 `.xlsb` 二进制文件。
-   **`pydantic`**: 用于定义和验证数据模型，例如 `TableModel`。

## 6. 贡献代码

我们欢迎任何形式的贡献。如果您希望为项目贡献代码，请遵循以下流程：

1.  Fork 本仓库。
2.  基于 `main` 分支创建一个新的功能分支 (`git checkout -b feature/your-feature-name`)。
3.  提交您的代码修改。请确保您的代码遵循项目现有的编码风格。
4.  为您的新功能添加相应的测试用例。
5.  确保所有测试都已通过。
6.  发起一个 Pull Request，并详细描述您的修改内容。