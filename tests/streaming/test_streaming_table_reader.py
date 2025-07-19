
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
