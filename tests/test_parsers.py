import pytest
from pathlib import Path
from src.models.table_model import Sheet, Row, Cell
from src.parsers.csv_parser import CsvParser
from src.parsers.xlsx_parser import XlsxParser
from src.parsers.factory import ParserFactory, UnsupportedFileType

@pytest.fixture
def sample_csv_path() -> Path:
    return Path(__file__).parent / "data" / "sample.csv"

@pytest.fixture
def sample_xlsx_path() -> Path:
    return Path(__file__).parent / "data" / "sample.xlsx"

def test_csv_parser_success(sample_csv_path: Path):
    """
    Tests that the CsvParser can successfully parse a simple CSV file.
    """
    # Arrange
    parser = CsvParser()
    expected_sheet = Sheet(
        name="sample",
        rows=[
            Row(cells=[Cell(value="header1"), Cell(value="header2"), Cell(value="header3")]),
            Row(cells=[Cell(value="value1_1"), Cell(value="value1_2"), Cell(value="value1_3")]),
            Row(cells=[Cell(value="value2_1"), Cell(value="value2_2"), Cell(value="value2_3")]),
            Row(cells=[Cell(value="value3_1"), Cell(value="value3_2"), Cell(value="value3_3")]),
        ]
    )

    # Act
    result_sheet = parser.parse(str(sample_csv_path))

    # Assert
    assert isinstance(result_sheet, Sheet)
    assert result_sheet.name == "sample"
    assert len(result_sheet.rows) == 4
    
    # Check headers
    assert [cell.value for cell in result_sheet.rows[0].cells] == ["header1", "header2", "header3"]
    
    # Check data rows
    assert [cell.value for cell in result_sheet.rows[1].cells] == ["value1_1", "value1_2", "value1_3"]
    assert [cell.value for cell in result_sheet.rows[2].cells] == ["value2_1", "value2_2", "value2_3"]
    assert [cell.value for cell in result_sheet.rows[3].cells] == ["value3_1", "value3_2", "value3_3"]

    # Deep comparison
    assert result_sheet == expected_sheet

def test_xlsx_parser_success(sample_xlsx_path: Path):
    """
    Tests that the XlsxParser can successfully parse a simple XLSX file.
    """
    # Arrange
    parser = XlsxParser()
    expected_sheet = Sheet(
        name="Sheet1",
        rows=[
            Row(cells=[Cell(value="ID"), Cell(value="Name"), Cell(value="Value")]),
            Row(cells=[Cell(value=1), Cell(value="Alice"), Cell(value=100)]),
            Row(cells=[Cell(value=2), Cell(value="Bob"), Cell(value=200)]),
        ]
    )

    # Act
    result_sheet = parser.parse(str(sample_xlsx_path))

    # Assert
    assert isinstance(result_sheet, Sheet)
    assert result_sheet.name == "Sheet1"
    assert len(result_sheet.rows) == 3
    
    # Check headers
    assert [cell.value for cell in result_sheet.rows[0].cells] == ["ID", "Name", "Value"]
    
    # Check data rows
    assert [cell.value for cell in result_sheet.rows[1].cells] == [1, "Alice", 100]
    assert [cell.value for cell in result_sheet.rows[2].cells] == [2, "Bob", 200]

    # Deep comparison
    assert result_sheet == expected_sheet


class TestParserFactory:
    def test_get_parser_csv(self, sample_csv_path: Path):
        """
        Tests that the ParserFactory returns a CsvParser instance for .csv files.
        """
        # Act
        parser = ParserFactory.get_parser(str(sample_csv_path))
        # Assert
        assert isinstance(parser, CsvParser)

    def test_get_parser_xlsx(self, sample_xlsx_path: Path):
        """
        Tests that the ParserFactory returns an XlsxParser instance for .xlsx files.
        """
        # Act
        parser = ParserFactory.get_parser(str(sample_xlsx_path))
        # Assert
        assert isinstance(parser, XlsxParser)

    def test_get_parser_unsupported_file_type(self):
        """
        Tests that the ParserFactory raises an UnsupportedFileType error for unsupported file types.
        """
        # Arrange
        unsupported_file_path = "test.txt"
        # Act & Assert
        with pytest.raises(UnsupportedFileType):
            ParserFactory.get_parser(unsupported_file_path)
