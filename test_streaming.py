#!/usr/bin/env python3
"""
Simple test script to verify the streaming functionality implementation.
"""

import os
import sys
import csv
import tempfile
from pathlib import Path

# Add the src directory to Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from parsers.factory import ParserFactory
from models.table_model import Row, Cell


def create_test_csv(file_path: str, num_rows: int = 1000):
    """Create a test CSV file with specified number of rows."""
    with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        
        # Write header
        writer.writerow(['ID', 'Name', 'Email', 'Age', 'City'])
        
        # Write data rows
        for i in range(num_rows):
            writer.writerow([
                i + 1,
                f'User_{i + 1}',
                f'user{i + 1}@example.com',
                20 + (i % 50),
                f'City_{(i % 10) + 1}'
            ])


def test_csv_streaming():
    """Test CSV streaming functionality."""
    print("Testing CSV streaming...")
    
    # Create a temporary CSV file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as temp_file:
        temp_csv_path = temp_file.name
    
    try:
        # Create test data
        create_test_csv(temp_csv_path, 100)
        
        # Test if CSV format supports streaming
        supports_streaming = ParserFactory.supports_streaming(temp_csv_path)
        print(f"CSV supports streaming: {supports_streaming}")
        
        if supports_streaming:
            # Create lazy sheet
            lazy_sheet = ParserFactory.create_lazy_sheet(temp_csv_path)
            print(f"Created lazy sheet: {lazy_sheet.name}")
            print(f"Total rows: {lazy_sheet.get_total_rows()}")
            
            # Test streaming first 5 rows
            print("\nFirst 5 rows:")
            for i, row in enumerate(lazy_sheet.iter_rows(0, 5)):
                print(f"Row {i}: {[cell.value for cell in row.cells]}")
            
            # Test getting specific rows
            print(f"\nRow 10: {[cell.value for cell in lazy_sheet.get_row(10).cells]}")
            print(f"Row 50: {[cell.value for cell in lazy_sheet.get_row(50).cells]}")
            
            # Test slice access
            print(f"\nRows 20-25:")
            for i, row in enumerate(lazy_sheet.iter_rows(20, 5)):
                print(f"Row {20 + i}: {[cell.value for cell in row.cells]}")
        
        print("CSV streaming test completed successfully!")
        
    finally:
        # Clean up
        if os.path.exists(temp_csv_path):
            os.unlink(temp_csv_path)


def test_factory_info():
    """Test factory streaming information."""
    print("\n" + "="*50)
    print("Testing ParserFactory streaming information...")
    
    # Get all supported formats
    supported_formats = ParserFactory.get_supported_formats()
    print(f"Supported formats: {supported_formats}")
    
    # Get streaming formats
    streaming_formats = ParserFactory.get_streaming_formats()
    print(f"Streaming formats: {streaming_formats}")
    
    # Get detailed format info
    format_info = ParserFactory.get_format_info()
    print("\nFormat details:")
    for fmt, info in format_info.items():
        streaming_support = "✓" if info.get('supports_streaming', False) else "✗"
        print(f"  {fmt}: {info['name']} - Streaming: {streaming_support}")
        print(f"    Features: {', '.join(info['features'])}")


def main():
    """Main test function."""
    print("Testing Lazy-loading & Streaming Implementation")
    print("=" * 50)
    
    try:
        # Test factory information
        test_factory_info()
        
        # Test CSV streaming
        test_csv_streaming()
        
        print("\n" + "="*50)
        print("All tests completed successfully!")
        
    except Exception as e:
        print(f"Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
