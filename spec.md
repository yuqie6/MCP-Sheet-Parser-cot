# MCP Sheet Parser - 规范与伪代码

本文档定义了MCP Sheet Parser的核心模块、接口和业务逻辑的详细规范和伪代码。
它是后续开发和实现的蓝图。

## 1. 核心模块概览

- **SheetService**: 业务逻辑的核心协调器。
- **ParserFactory**: 根据文件扩展名动态选择和创建解析器实例。
- **BaseParser (抽象)**: 定义所有解析器必须实现的通用接口。
- **XlsxParser**: 负责解析`.xlsx`文件。
- **CsvParser**: 负责解析`.csv`文件。
- **HTMLConverter**: 将统一的`Sheet`数据模型转换为HTML字符串。
- **TableModel**: 定义了`Sheet`, `Row`, `Cell`, `Style`等核心数据结构。

## 2. SheetService - 业务逻辑层

CLASS SheetService:
    // 依赖注入解析器工厂和HTML转换器
    CONSTRUCTOR(parser_factory, html_converter):
        this.parser_factory = parser_factory
        this.html_converter = html_converter

    // --- 核心方法 ---

    METHOD convert_to_html(file_path):
        // 1. 验证文件路径
        IF file_path is not valid:
            RAISE FileNotFoundError("文件不存在")

        // 2. 获取文件扩展名
        file_extension = GET_EXTENSION(file_path)

        // 3. 使用工厂创建合适的解析器
        TRY:
            parser = this.parser_factory.get_parser(file_extension)
        CATCH UnsupportedFormatError as e:
            RAISE e // 向上抛出不支持的格式错误

        // 4. 调用解析器进行解析，返回统一的Sheet模型
        TRY:
            sheet_model = parser.parse(file_path)
        CATCH ParsingError as e:
            RAISE e // 向上抛出解析错误

        // 5. 调用转换器将Sheet模型转换为HTML
        html_content = this.html_converter.convert(sheet_model)

        // 6. 返回HTML内容
        RETURN html_content

    METHOD convert_to_html_file(file_path, output_path):
        // 调用 convert_to_html 获取内容
        html_content = this.convert_to_html(file_path)

        // 将内容写入指定的输出文件
        WRITE_FILE(output_path, html_content)

        // 返回成功信息和输出路径
        RETURN { "status": "success", "output_path": output_path }

    METHOD get_summary(file_path):
        // ... (逻辑类似于 convert_to_html, 但返回摘要信息)
        // 1. 获取解析器并解析
        parser = this.parser_factory.get_parser(GET_EXTENSION(file_path))
        sheet_model = parser.parse(file_path)

        // 2. 生成摘要信息
        summary = {
            "sheet_name": sheet_model.name,
            "total_rows": LENGTH(sheet_model.rows),
            "total_columns": LENGTH(sheet_model.rows[0].cells) IF sheet_model.rows ELSE 0,
            "merged_cells_count": LENGTH(sheet_model.merged_cells),
            "sample_data": sheet_model.rows[:5] // 提取前5行作为样本
        }

        // 3. 返回摘要
        RETURN summary


## 3. Parser - 解析器层

### 3.1 ParserFactory

CLASS ParserFactory:
    // 注册所有可用的解析器
    CONSTRUCTOR():
        this.parsers = {
            "xlsx": XlsxParser(),
            "csv": CsvParser()
            // 未来可扩展其他格式, e.g., "xls": XlsParser()
        }

    METHOD get_parser(file_extension):
        // 转换为小写以兼容大小写不一的扩展名
        extension = file_extension.lower()

        // 查找并返回对应的解析器实例
        parser = this.parsers.get(extension)

        IF parser is None:
            RAISE UnsupportedFormatError(f"不支持的文件格式: {extension}")

        RETURN parser

### 3.2 BaseParser (抽象接口)

ABSTRACT CLASS BaseParser:
    // 定义所有解析器都必须实现的方法
    ABSTRACT METHOD parse(file_path) -> Sheet

### 3.3 XlsxParser

CLASS XlsxParser IMPLEMENTS BaseParser:
    METHOD parse(file_path) -> Sheet:
        // 使用 openpyxl 库加载Excel文件
        workbook = LOAD_WORKBOOK(file_path)
        worksheet = workbook.active // 默认处理第一个工作表

        // 创建Sheet数据模型
        rows = []
        FOR each row_data in worksheet.iter_rows():
            cells = []
            FOR each cell_data in row_data:
                // 提取单元格值和基本样式
                cell_value = cell_data.value
                cell_style = this._extract_style(cell_data)
                cell = Cell(value=cell_value, style=cell_style)
                cells.append(cell)
            rows.append(Row(cells=cells))

        // 提取合并单元格信息
        merged_cells_info = worksheet.merged_cells

        RETURN Sheet(name=worksheet.title, rows=rows, merged_cells=merged_cells_info)

    PRIVATE METHOD _extract_style(cell_data) -> Style:
        // 从cell_data中提取字体、颜色等样式信息
        // ... 实现细节 ...
        RETURN Style(...) 

### 3.4 CsvParser

CLASS CsvParser IMPLEMENTS BaseParser:
    METHOD parse(file_path) -> Sheet:
        // 使用 csv 库读取CSV文件
        rows = []
        WITH open(file_path, mode='r', encoding='utf-8') as csvfile:
            csv_reader = CSV_READER(csvfile)
            FOR each row_data in csv_reader:
                cells = [Cell(value=item) for item in row_data]
                rows.append(Row(cells=cells))

        // CSV没有复杂的样式或合并单元格
        RETURN Sheet(name=GET_FILENAME(file_path), rows=rows)


## 4. HTMLConverter - HTML转换器

CLASS HTMLConverter:
    // 使用Jinja2等模板引擎进行渲染
    CONSTRUCTOR(template_engine):
        this.template_engine = template_engine

    METHOD convert(sheet_model) -> String:
        // 准备模板所需的数据上下文
        context = {
            "title": sheet_model.name,
            "rows": sheet_model.rows,
            "merged_cells": sheet_model.merged_cells,
            "styles": this._generate_css(sheet_model) // 生成CSS样式
        }

        // 加载HTML模板并渲染
        template = this.template_engine.get_template("table_template.html")
        html_output = template.render(context)

        RETURN html_output

    PRIVATE METHOD _generate_css(sheet_model) -> String:
        // 遍历Sheet模型，提取所有唯一样式
        // 将Style对象转换为CSS类
        // e.g., Style(bold=True, font_color="#FF0000") -> .style_1 { font-weight: bold; color: #FF0000; }
        // ... 实现细节 ...
        RETURN css_string

## 5. TableModel - 核心数据模型

// 使用dataclass或类似结构定义，清晰且不可变

DATA CLASS Style:
    bold: Boolean = False
    italic: Boolean = False
    font_color: String = "#000000"
    background_color: String = "#FFFFFF"
    // ... 可根据需要扩展更多样式属性

DATA CLASS Cell:
    value: Any
    style: Optional[Style] = None
    // 单元格的原始行列信息，用于处理合并等场景
    row_span: Integer = 1 
    col_span: Integer = 1

DATA CLASS Row:
    cells: List[Cell]

DATA CLASS Sheet:
    name: String
    rows: List[Row]
    // 存储合并单元格的范围，e.g., ["A1:B2", "C3:C5"]
    merged_cells: List[String]
