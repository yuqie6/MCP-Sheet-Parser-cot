"""
项目常量定义

集中管理项目中使用的所有常量，避免魔法数字和字符串。
"""

# 文件格式相关常量
class FileFormats:
    """支持的文件格式常量。"""
    CSV = "csv"
    XLSX = "xlsx"
    XLS = "xls"
    XLSB = "xlsb"
    XLSM = "xlsm"
    
    # 所有支持的格式
    ALL_FORMATS = [CSV, XLSX, XLS, XLSB, XLSM]
    
    # 格式描述
    DESCRIPTIONS = {
        CSV: "逗号分隔值文件",
        XLSX: "Excel 2007+ 格式",
        XLS: "Excel 97-2003 格式",
        XLSB: "Excel 二进制格式",
        XLSM: "Excel 宏启用格式"
    }

# 样式相关常量
class StyleConstants:
    """样式处理相关常量。"""
    
    # 默认颜色
    DEFAULT_FONT_COLOR = "#000000"
    DEFAULT_BACKGROUND_COLOR = "#FFFFFF"
    DEFAULT_BORDER_COLOR = "#000000"
    
    # 字体大小限制
    MIN_FONT_SIZE_PT = 6
    MAX_FONT_SIZE_PT = 72
    DEFAULT_FONT_SIZE_PT = 12
    
    # 边框样式映射
    BORDER_STYLES = {
        0: None,
        1: "thin",
        2: "medium",
        3: "dashed",
        4: "dotted",
        5: "thick",
        6: "double",
        7: "hair",
        8: "medium_dashed",
        9: "dash_dot",
        10: "medium_dash_dot",
        11: "dash_dot_dot",
        12: "medium_dash_dot_dot",
        13: "slant_dash_dot"
    }
    
    # 字体回退映射
    FONT_FALLBACKS = {
        "chinese": '"Microsoft YaHei", "SimHei", "SimSun", sans-serif',
        "monospace": '"Courier New", "Consolas", "Monaco", monospace',
        "serif": '"Times New Roman", "Georgia", serif',
        "sans_serif": '"Arial", "Helvetica", sans-serif'
    }
    
    # 颜色名称映射
    COLOR_NAMES = {
        "BLACK": "#000000",
        "WHITE": "#FFFFFF",
        "RED": "#FF0000",
        "GREEN": "#00FF00",
        "BLUE": "#0000FF",
        "YELLOW": "#FFFF00",
        "CYAN": "#00FFFF",
        "MAGENTA": "#FF00FF",
        "GRAY": "#808080",
        "GREY": "#808080"
    }


# HTML转换相关常量
class HTMLConstants:
    """HTML转换相关常量。"""
    
    # HTML模板
    HTML_DOCTYPE = "<!DOCTYPE html>"
    HTML_LANG = "zh-CN"
    
    # CSS类名前缀
    CSS_CLASS_PREFIX = "style_"
    
    # 表格相关
    TABLE_ROLE = "table"
    HEADER_TAG = "th"
    DATA_TAG = "td"
    
    # 默认CSS样式
    DEFAULT_TABLE_CSS = """
        table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
            color: #000000 !important;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
            color: #000000 !important;
        }
        th {
            background-color: #f2f2f2;
            font-weight: bold;
            color: #000000 !important;
        }
        body {
            color: #000000 !important;
            background-color: #ffffff !important;
        }
    """


# 缓存相关常量
class CacheConstants:
    """缓存相关常量。"""
    
    # 缓存类型
    MEMORY_CACHE = "memory"
    DISK_CACHE = "disk"
    
    # 缓存格式
    PICKLE_FORMAT = "pickle"
    PARQUET_FORMAT = "parquet"
    
    # 缓存目录名
    CACHE_DIR_NAME = "mcp-sheet-parser"
    
    # 缓存文件扩展名
    CACHE_FILE_EXTENSIONS = {
        PICKLE_FORMAT: ".pkl",
        PARQUET_FORMAT: ".parquet"
    }


# 日志相关常量
class LogConstants:
    """日志相关常量。"""
    
    # 日志级别
    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"
    
    # 日志格式
    DEFAULT_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    DETAILED_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s"
    
    # 日志文件
    DEFAULT_LOG_FILE = "mcp_sheet_parser.log"


# 错误代码常量
class ErrorCodes:
    """错误代码常量。"""
    
    # 文件相关错误
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    FILE_ACCESS_DENIED = "FILE_ACCESS_DENIED"
    UNSUPPORTED_FILE_TYPE = "UNSUPPORTED_FILE_TYPE"
    CORRUPTED_FILE = "CORRUPTED_FILE"
    
    # 解析相关错误
    SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
    INVALID_RANGE = "INVALID_RANGE"
    STYLE_EXTRACTION_ERROR = "STYLE_EXTRACTION_ERROR"
    
    # 转换相关错误
    HTML_CONVERSION_ERROR = "HTML_CONVERSION_ERROR"
    
    # 验证相关错误
    VALIDATION_ERROR = "VALIDATION_ERROR"
    
    # 资源相关错误
    MEMORY_LIMIT_EXCEEDED = "MEMORY_LIMIT_EXCEEDED"
    FILE_SIZE_LIMIT_EXCEEDED = "FILE_SIZE_LIMIT_EXCEEDED"
    OPERATION_TIMEOUT = "OPERATION_TIMEOUT"
    
    # 缓存相关错误
    CACHE_ERROR = "CACHE_ERROR"
    
    # 配置相关错误
    CONFIGURATION_ERROR = "CONFIGURATION_ERROR"


# 正则表达式常量
class RegexPatterns:
    """正则表达式模式常量。"""
    
    # 单元格范围模式
    CELL_RANGE = r'^([A-Z]+\d+)(?::([A-Z]+\d+))?$'
    SINGLE_CELL = r'^[A-Z]+\d+$'
    
    # 颜色模式
    HEX_COLOR_6 = r'^#[0-9A-F]{6}$'
    HEX_COLOR_3 = r'^#[0-9A-F]{3}$'
    HEX_COLOR_NO_HASH_6 = r'^[0-9A-F]{6}$'
    HEX_COLOR_NO_HASH_3 = r'^[0-9A-F]{3}$'
    
    # 危险路径模式
    DANGEROUS_PATHS = [
        r'\.\./',
        r'\.\.\\',
        r'/etc/',
        r'/proc/',
        r'/sys/',
        r'C:\\Windows\\',
        r'C:\\Program Files\\'
    ]
    
    # 边框样式模式
    BORDER_STYLE_PATTERN = r'(\d+(?:\.\d+)?)(px|pt|em|rem)?\s*(solid|dashed|dotted|double|groove|ridge|inset|outset)?'


# 数字格式常量
class NumberFormats:
    """数字格式常量。"""
    
    # 基本格式
    GENERAL = "General"
    INTEGER = "0"
    DECIMAL_1 = "0.0"
    DECIMAL_2 = "0.00"
    
    # 千分位格式
    THOUSAND_INTEGER = "#,##0"
    THOUSAND_DECIMAL_1 = "#,##0.0"
    THOUSAND_DECIMAL_2 = "#,##0.00"
    
    # 百分比格式
    PERCENT_INTEGER = "0%"
    PERCENT_DECIMAL_1 = "0.0%"
    PERCENT_DECIMAL_2 = "0.00%"
    
    # 日期格式
    DATE_FORMATS = {
        "yyyy-mm-dd": "%Y-%m-%d",
        "mm/dd/yyyy": "%m/%d/%Y",
        "dd/mm/yyyy": "%d/%m/%Y",
        "yyyy/mm/dd": "%Y/%m/%d",
        "mm-dd-yyyy": "%m-%d-%Y",
        "dd-mm-yyyy": "%d-%m-%Y"
    }


# MCP协议相关常量
class MCPConstants:
    """MCP协议相关常量。"""
    
    # 工具名称
    TOOL_PARSE_SHEET = "parse_sheet"
    TOOL_CONVERT_TO_HTML = "convert_to_html"
    TOOL_APPLY_CHANGES = "apply_changes"
    
    # 服务器信息
    SERVER_NAME = "mcp-sheet-parser"
    SERVER_VERSION = "0.1.0"
    
    # 协议版本
    PROTOCOL_VERSION = "2025-06-18"


# 图表处理相关常量
class ChartConstants:
    """图表数据提取和处理相关常量。"""

    # 字符串处理限制
    MAX_STRING_REPRESENTATION_LENGTH = 100
    MAX_TITLE_LENGTH = 200

    # 图表元素默认值
    DEFAULT_PIE_SLICE_COUNT = 3
    MAX_DATA_LABEL_ITEMS = 50
    DEFAULT_CHART_WIDTH = 600
    DEFAULT_CHART_HEIGHT = 400

    # 颜色相关
    DEFAULT_COLOR_PALETTE_SIZE = 10
    MAX_COLOR_VARIANTS = 20
