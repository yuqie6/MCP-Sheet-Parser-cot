"""
Parses cell range strings like "A1:D10" or "A1".
"""
import re

def parse_range_string(range_string: str) -> tuple[int, int, int, int]:
    """
    Parses a range string, e.g., "A1:D10" or "A1".

    Returns:
        (start_row, start_col, end_row, end_col) - 0-based indices
    """
    range_string = range_string.strip().upper()

    single_cell_pattern = r'^([A-Z]+)(\d+)$'
    range_pattern = r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$'

    def col_to_num(col_str):
        """Converts column letters to numbers (A=0, B=1, ...)."""
        result = 0
        for char in col_str:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1

    range_match = re.match(range_pattern, range_string)
    if range_match:
        start_col_str, start_row_str, end_col_str, end_row_str = range_match.groups()
        start_row = int(start_row_str) - 1
        start_col = col_to_num(start_col_str)
        end_row = int(end_row_str) - 1
        end_col = col_to_num(end_col_str)
        return start_row, start_col, end_row, end_col

    single_match = re.match(single_cell_pattern, range_string)
    if single_match:
        col_str, row_str = single_match.groups()
        row = int(row_str) - 1
        col = col_to_num(col_str)
        return row, col, row, col

    raise ValueError(f"Invalid range format: {range_string}. Supported formats: A1 or A1:D10")
