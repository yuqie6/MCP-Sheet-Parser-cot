#!/usr/bin/env python3
"""测试修复后的图表生成"""

import sys
import os
from pathlib import Path

from src.parsers.xlsx_parser import XlsxParser
from src.converters.html_converter import HTMLConverter

def test_chart_generation():
    """测试图表生成"""
    file_path = "销售费用分析表.xlsx"
    
    if not os.path.exists(file_path):
        print(f"错误：文件 {file_path} 不存在")
        return
    
    print("=== 测试修复后的图表生成 ===\n")
    
    # 1. 解析文件
    print("1. 解析Excel文件...")
    try:
        parser = XlsxParser()
        sheets = parser.parse(file_path)
        print(f"解析成功：{len(sheets)} 个工作表")
        
        # 查看图表信息
        for sheet in sheets:
            print(f"  工作表: {sheet.name}")
            print(f"  行数: {len(sheet.rows)}")
            print(f"  列数: {len(sheet.rows[0].cells) if sheet.rows else 0}")
            print(f"  图表数: {len(sheet.charts)}")
            
            # 显示图表详细信息
            for i, chart in enumerate(sheet.charts):
                print(f"    图表 {i+1}: {chart.name} ({chart.type}) 位置={chart.anchor}")
                print(f"      图表数据: {bool(chart.chart_data)} SVG数据: {bool(chart.svg_data)}")
                
    except Exception as e:
        print(f"解析错误: {e}")
        return
    
    # 2. 生成HTML
    print("\n2. 生成HTML...")
    try:
        converter = HTMLConverter()
        html_results = converter.convert_to_files(sheets, "test_chart_output.html")
        
        for result in html_results:
            if result['status'] == 'success':
                print(f"HTML生成成功：{result['sheet_name']}")
                print(f"  输出路径: {result['output_path']}")
                print(f"  文件大小: {result['file_size_kb']} KB")
                print(f"  包含样式: {result['has_styles']}")
                print(f"  包含合并单元格: {result['has_merged_cells']}")
            else:
                print(f"HTML生成失败：{result['sheet_name']} - {result['error']}")
                
    except Exception as e:
        print(f"HTML生成错误: {e}")
        return
    
    print("\n测试完成！请查看生成的HTML文件以验证图表显示效果。")

if __name__ == "__main__":
    test_chart_generation()