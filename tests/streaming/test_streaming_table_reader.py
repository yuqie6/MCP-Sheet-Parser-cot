
import pytest
from unittest.mock import MagicMock, patch
from src.streaming.streaming_table_reader import StreamingTableReader, ChunkFilter, StreamingChunk
from src.models.table_model import Row, Cell, LazySheet

@pytest.fixture
def mock_parser():
    """Fixture for a mocked parser that supports streaming."""
    parser = MagicMock()
    parser.supports_streaming.return_value = True
    lazy_sheet = MagicMock(spec=LazySheet)
    lazy_sheet.get_total_rows.return_value = 100
    header_row = Row(cells=[Cell(value="Header1"), Cell(value="Header2")])
    lazy_sheet.get_row.return_value = header_row
    lazy_sheet.iter_rows.return_value = [Row(cells=[Cell(value=f"d{i}")]) for i in range(10)]
    parser.create_lazy_sheet.return_value = lazy_sheet
    return parser

@patch('src.streaming.streaming_table_reader.ParserFactory.get_parser')
def test_reader_initialization_streaming(mock_get_parser, mock_parser):
    """Test reader initialization with a streaming parser."""
    mock_get_parser.return_value = mock_parser
    reader = StreamingTableReader("dummy.xlsx")
    assert reader._lazy_sheet is not None
    assert reader._regular_sheet is None

def test_iter_chunks(mock_parser):
    """Test iterating through chunks."""
    with patch('src.streaming.streaming_table_reader.ParserFactory.get_parser', return_value=mock_parser):
        reader = StreamingTableReader("dummy.xlsx")
        chunks = list(reader.iter_chunks(rows=10))
        assert len(chunks) > 0
        assert isinstance(chunks[0], StreamingChunk)
        assert len(chunks[0].rows) == 10

def test_iter_chunks_with_filter(mock_parser):
    """Test iterating with a filter."""
    with patch('src.streaming.streaming_table_reader.ParserFactory.get_parser', return_value=mock_parser):
        reader = StreamingTableReader("dummy.xlsx")
        filter_config = ChunkFilter(columns=["Header1"])
        chunks = list(reader.iter_chunks(rows=10, filter_config=filter_config))
        assert len(chunks[0].headers) == 1
        assert chunks[0].headers[0] == "Header1"

def test_parse_range_filter():
    """Test _parse_range_filter."""
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())
    result = reader._parse_range_filter("A1:B10")
    assert result['start_row'] == 0
    assert result['max_rows'] == 10
    assert result['column_indices'] == [0, 1]

def test_col_to_index():
    """Test _col_to_index."""
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())
    assert reader._col_to_index('A') == 0
    assert reader._col_to_index('Z') == 25
    assert reader._col_to_index('AA') == 26

# === TDD测试：提升StreamingTableReader覆盖率 ===

@patch('src.streaming.streaming_table_reader.ParserFactory.get_parser')
def test_reader_initialization_non_streaming(mock_get_parser):
    """
    TDD测试：StreamingTableReader应该处理不支持流式读取的解析器

    这个测试覆盖第74-76行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    non_streaming_parser = MagicMock()
    non_streaming_parser.supports_streaming.return_value = False
    non_streaming_parser.parse.return_value = [MagicMock()]
    mock_get_parser.return_value = non_streaming_parser

    reader = StreamingTableReader("dummy.xlsx")

    # 应该使用常规解析而不是流式解析
    assert reader._lazy_sheet is None
    assert reader._regular_sheet is not None

def test_iter_chunks_with_non_streaming_fallback():
    """
    TDD测试：iter_chunks应该在非流式模式下正确工作

    这个测试覆盖第90行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    non_streaming_parser = MagicMock()
    non_streaming_parser.supports_streaming.return_value = False

    # 创建模拟的常规工作表
    regular_sheet = MagicMock()
    regular_sheet.rows = [
        Row(cells=[Cell(value="Header1"), Cell(value="Header2")]),
        Row(cells=[Cell(value="Data1"), Cell(value="Data2")]),
        Row(cells=[Cell(value="Data3"), Cell(value="Data4")])
    ]
    non_streaming_parser.parse.return_value = [regular_sheet]

    with patch('src.streaming.streaming_table_reader.ParserFactory.get_parser', return_value=non_streaming_parser):
        reader = StreamingTableReader("dummy.xlsx")
        chunks = list(reader.iter_chunks(rows=2))

        # 应该能够生成块
        assert len(chunks) > 0
        assert isinstance(chunks[0], StreamingChunk)

def test_iter_chunks_with_range_filter():
    """
    TDD测试：iter_chunks应该处理范围过滤器

    这个测试覆盖第94-97行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    mock_parser = MagicMock()
    mock_parser.supports_streaming.return_value = True

    lazy_sheet = MagicMock(spec=LazySheet)
    lazy_sheet.get_total_rows.return_value = 100
    lazy_sheet.get_row.return_value = Row(cells=[Cell(value="Header1"), Cell(value="Header2")])
    lazy_sheet.iter_rows.return_value = [
        Row(cells=[Cell(value="Data1"), Cell(value="Data2")]),
        Row(cells=[Cell(value="Data3"), Cell(value="Data4")])
    ]
    mock_parser.create_lazy_sheet.return_value = lazy_sheet

    with patch('src.streaming.streaming_table_reader.ParserFactory.get_parser', return_value=mock_parser):
        reader = StreamingTableReader("dummy.xlsx")

        # 使用范围过滤器
        chunks = list(reader.iter_chunks(rows=10, range_filter="A1:B5"))

        # 应该调用_parse_range_filter
        assert len(chunks) > 0

def test_parse_range_filter_with_invalid_range():
    """
    TDD测试：_parse_range_filter应该处理无效的范围

    这个测试覆盖第113行的异常处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())

    # 测试无效的范围格式
    result = reader._parse_range_filter("INVALID_RANGE")

    # 应该返回None或默认值
    assert result is None

def test_parse_range_filter_with_single_cell():
    """
    TDD测试：_parse_range_filter应该处理单个单元格引用

    这个测试覆盖单个单元格的处理逻辑
    """
    # 🔴 红阶段：编写测试描述期望的行为
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())

    result = reader._parse_range_filter("B5")

    # 应该正确解析单个单元格
    assert result['start_row'] == 4  # B5的行索引
    assert result['max_rows'] == 1
    assert result['column_indices'] == [1]  # B列的索引

def test_apply_chunk_filter_with_column_filter():
    """
    TDD测试：_apply_chunk_filter应该正确应用列过滤器

    这个测试覆盖第127行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())

    # 创建测试数据
    headers = ["Col1", "Col2", "Col3"]
    rows = [
        Row(cells=[Cell(value="A1"), Cell(value="B1"), Cell(value="C1")]),
        Row(cells=[Cell(value="A2"), Cell(value="B2"), Cell(value="C2")])
    ]

    filter_config = ChunkFilter(columns=["Col1", "Col3"])

    filtered_headers, filtered_rows = reader._apply_chunk_filter(headers, rows, filter_config)

    # 应该只保留指定的列
    assert filtered_headers == ["Col1", "Col3"]
    assert len(filtered_rows[0].cells) == 2
    assert filtered_rows[0].cells[0].value == "A1"
    assert filtered_rows[0].cells[1].value == "C1"

def test_apply_chunk_filter_with_row_filter():
    """
    TDD测试：_apply_chunk_filter应该正确应用行过滤器

    这个测试覆盖第130行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())

    # 创建测试数据
    headers = ["Col1", "Col2"]
    rows = [
        Row(cells=[Cell(value="A1"), Cell(value="B1")]),
        Row(cells=[Cell(value="A2"), Cell(value="B2")]),
        Row(cells=[Cell(value="A3"), Cell(value="B3")])
    ]

    # 创建行过滤器函数
    def row_filter(row_data):
        return row_data[0] != "A2"  # 过滤掉第二行

    filter_config = ChunkFilter(row_filter=row_filter)

    filtered_headers, filtered_rows = reader._apply_chunk_filter(headers, rows, filter_config)

    # 应该过滤掉第二行
    assert len(filtered_rows) == 2
    assert filtered_rows[0].cells[0].value == "A1"
    assert filtered_rows[1].cells[0].value == "A3"

def test_apply_chunk_filter_with_no_filter():
    """
    TDD测试：_apply_chunk_filter应该在没有过滤器时返回原始数据

    这个测试确保方法在没有过滤器时的行为
    """
    # 🔴 红阶段：编写测试描述期望的行为
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())

    headers = ["Col1", "Col2"]
    rows = [Row(cells=[Cell(value="A1"), Cell(value="B1")])]

    filtered_headers, filtered_rows = reader._apply_chunk_filter(headers, rows, None)

    # 应该返回原始数据
    assert filtered_headers == headers
    assert filtered_rows == rows

def test_get_column_indices_with_valid_columns():
    """
    TDD测试：_get_column_indices应该正确映射列名到索引

    这个测试覆盖第162-163行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())

    headers = ["Name", "Age", "City", "Country"]
    columns = ["Age", "Country"]

    indices = reader._get_column_indices(headers, columns)

    # 应该返回正确的索引
    assert indices == [1, 3]

def test_get_column_indices_with_invalid_columns():
    """
    TDD测试：_get_column_indices应该处理无效的列名

    这个测试覆盖第170行的异常处理代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())

    headers = ["Name", "Age", "City"]
    columns = ["Age", "InvalidColumn"]

    indices = reader._get_column_indices(headers, columns)

    # 应该只返回有效列的索引
    assert indices == [1]

def test_filter_row_by_indices():
    """
    TDD测试：_filter_row_by_indices应该正确过滤行数据

    这个测试覆盖第178-181行的代码路径
    """
    # 🔴 红阶段：编写测试描述期望的行为
    reader = StreamingTableReader("dummy.xlsx", parser=MagicMock())

    row = Row(cells=[
        Cell(value="A"), Cell(value="B"),
        Cell(value="C"), Cell(value="D")
    ])
    indices = [0, 2, 3]

    filtered_row = reader._filter_row_by_indices(row, indices)

    # 应该只保留指定索引的单元格
    assert len(filtered_row.cells) == 3
    assert filtered_row.cells[0].value == "A"
    assert filtered_row.cells[1].value == "C"
    assert filtered_row.cells[2].value == "D"
