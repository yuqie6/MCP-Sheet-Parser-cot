import pytest
from src.models.table_model import Sheet, Row, Cell
from src.converters.html_converter import HTMLConverter

def test_html_converter():
    # 1. Arrange: Create a simple Sheet object
    sheet = Sheet(
        name="Test Sheet",
        rows=[
            Row(cells=[Cell(value='A1'), Cell(value='B1')]),
            Row(cells=[Cell(value='A2'), Cell(value='B2')])
        ]
    )

    # 2. Act: Convert the Sheet to an HTML string
    converter = HTMLConverter()
    html_output = converter.convert(sheet)

    # 3. Assert: Check if the output is a valid HTML table
    assert '<table>' in html_output
    assert '</table>' in html_output
    expected_row1 = '<tr><tdstyle="color:#000000;background-color:#FFFFFF;font-weight:normal;font-style:normal;"colspan="1"rowspan="1">A1</td><tdstyle="color:#000000;background-color:#FFFFFF;font-weight:normal;font-style:normal;"colspan="1"rowspan="1">B1</td></tr>'
    expected_row2 = '<tr><tdstyle="color:#000000;background-color:#FFFFFF;font-weight:normal;font-style:normal;"colspan="1"rowspan="1">A2</td><tdstyle="color:#000000;background-color:#FFFFFF;font-weight:normal;font-style:normal;"colspan="1"rowspan="1">B2</td></tr>'

    processed_html = ''.join(html_output.split())

    assert expected_row1 in processed_html
    assert expected_row2 in processed_html
