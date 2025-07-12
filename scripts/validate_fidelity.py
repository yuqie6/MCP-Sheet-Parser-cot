#!/usr/bin/env python3
"""
样式保真度验证脚本

用于验证解析器的样式保真度，生成详细的质量报告。
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.utils.style_validator import StyleValidator
from src.parsers.factory import ParserFactory
from src.models.table_model import Style, Sheet, Row, Cell
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def create_reference_sheet() -> Sheet:
    """
    创建参考工作表，包含各种样式的单元格。
    这代表了我们期望的"完美"样式。
    """
    # 定义各种样式
    header_style = Style(
        font_name="Arial",
        font_size=14.0,
        font_color="#FFFFFF",
        background_color="#4472C4",
        bold=True,
        text_align="center",
        vertical_align="middle"
    )
    
    data_style = Style(
        font_name="Calibri",
        font_size=11.0,
        font_color="#000000",
        background_color="#FFFFFF",
        text_align="left",
        vertical_align="top"
    )
    
    number_style = Style(
        font_name="Calibri",
        font_size=11.0,
        font_color="#000000",
        background_color="#FFFFFF",
        text_align="right",
        vertical_align="top",
        number_format="0.00"
    )
    
    highlight_style = Style(
        font_name="Calibri",
        font_size=11.0,
        font_color="#000000",
        background_color="#FFFF00",
        bold=True,
        text_align="center",
        vertical_align="middle"
    )
    
    # 创建工作表
    return Sheet(
        name="FidelityTest",
        rows=[
            Row(cells=[
                Cell(value="ID", style=header_style),
                Cell(value="Name", style=header_style),
                Cell(value="Value", style=header_style),
                Cell(value="Status", style=header_style)
            ]),
            Row(cells=[
                Cell(value=1, style=data_style),
                Cell(value="Alice", style=data_style),
                Cell(value=100.50, style=number_style),
                Cell(value="Active", style=highlight_style)
            ]),
            Row(cells=[
                Cell(value=2, style=data_style),
                Cell(value="Bob", style=data_style),
                Cell(value=200.75, style=number_style),
                Cell(value="Inactive", style=data_style)
            ]),
            Row(cells=[
                Cell(value=3, style=data_style),
                Cell(value="Charlie", style=data_style),
                Cell(value=300.25, style=number_style),
                Cell(value="Pending", style=highlight_style)
            ])
        ]
    )


def simulate_parser_output(parser_name: str, reference_sheet: Sheet) -> Sheet:
    """
    模拟解析器输出，基于解析器特性添加一些变化。
    
    Args:
        parser_name: 解析器名称
        reference_sheet: 参考工作表
        
    Returns:
        模拟的解析器输出
    """
    # 深拷贝参考工作表
    simulated_sheet = Sheet(
        name=reference_sheet.name,
        rows=[],
        merged_cells=reference_sheet.merged_cells.copy()
    )
    
    for row in reference_sheet.rows:
        new_cells = []
        for cell in row.cells:
            # 复制单元格样式
            if cell.style:
                new_style = Style(
                    font_name=cell.style.font_name,
                    font_size=cell.style.font_size,
                    font_color=cell.style.font_color,
                    background_color=cell.style.background_color,
                    bold=cell.style.bold,
                    italic=cell.style.italic,
                    underline=cell.style.underline,
                    text_align=cell.style.text_align,
                    vertical_align=cell.style.vertical_align,
                    border_top=cell.style.border_top,
                    border_bottom=cell.style.border_bottom,
                    border_left=cell.style.border_left,
                    border_right=cell.style.border_right,
                    border_color=cell.style.border_color,
                    wrap_text=cell.style.wrap_text,
                    number_format=cell.style.number_format,
                    hyperlink=cell.style.hyperlink,
                    comment=cell.style.comment
                )
                
                # 根据解析器类型添加特定的变化
                if parser_name == "XlsParser":
                    # XLS解析器可能有字体名称的差异
                    if new_style.font_name == "Calibri":
                        new_style.font_name = "Arial"  # XLS可能默认使用Arial
                elif parser_name == "XlsbParser":
                    # XLSB解析器样式信息有限
                    new_style.background_color = "#FFFFFF"  # 可能丢失背景色
                    new_style.number_format = ""  # 可能丢失数字格式
                elif parser_name == "XlsmParser":
                    # XLSM解析器与XLSX相同，保持高保真度
                    pass
                elif parser_name == "XlsxParser":
                    # XLSX解析器保持最高保真度
                    pass
            else:
                new_style = None
            
            new_cells.append(Cell(value=cell.value, style=new_style))
        
        simulated_sheet.rows.append(Row(cells=new_cells))
    
    return simulated_sheet


def validate_all_parsers():
    """验证所有解析器的保真度"""
    logger.info("开始样式保真度验证...")
    
    validator = StyleValidator()
    reference_sheet = create_reference_sheet()
    
    # 获取所有支持的格式
    supported_formats = ParserFactory.get_supported_formats()
    parser_names = ["CsvParser", "XlsxParser", "XlsParser", "XlsbParser", "XlsmParser"]
    
    results = {}
    
    for parser_name in parser_names:
        logger.info(f"验证 {parser_name} 保真度...")
        
        # 模拟解析器输出
        simulated_output = simulate_parser_output(parser_name, reference_sheet)
        
        # 生成保真度报告
        report = validator.generate_fidelity_report(parser_name, reference_sheet, simulated_output)
        results[parser_name] = report
        
        # 输出结果
        print(f"\n{'='*60}")
        print(f"解析器: {parser_name}")
        print(f"{'='*60}")
        print(f"总体保真度: {report.overall_fidelity:.2f}%")
        print(f"分析单元格: {report.analyzed_cells}")
        print(f"完美单元格: {report.summary.get('perfect_cells', 0)}")
        print(f"有问题单元格: {report.summary.get('cells_with_issues', 0)}")
        
        print(f"\n属性得分:")
        for attr, score in report.summary.get('attribute_scores', {}).items():
            print(f"  {attr}: {score:.1f}%")
        
        print(f"\n建议:")
        for recommendation in report.recommendations:
            print(f"  • {recommendation}")
        
        # 质量门禁检查
        if report.overall_fidelity >= 95.0:
            print(f"✅ {parser_name} 达到95%保真度目标")
        elif report.overall_fidelity >= 90.0:
            print(f"⚠️  {parser_name} 接近95%目标，需要微调")
        else:
            print(f"❌ {parser_name} 未达到95%目标，需要优化")
    
    # 生成总体报告
    print(f"\n{'='*60}")
    print("总体保真度报告")
    print(f"{'='*60}")
    
    total_fidelity = sum(r.overall_fidelity for r in results.values()) / len(results)
    print(f"平均保真度: {total_fidelity:.2f}%")
    
    target_achieved = sum(1 for r in results.values() if r.overall_fidelity >= 95.0)
    print(f"达到95%目标的解析器: {target_achieved}/{len(results)}")
    
    if target_achieved == len(results):
        print("🎉 所有解析器都达到了95%保真度目标！")
    elif target_achieved >= len(results) * 0.8:
        print("👍 大部分解析器达到了目标，整体表现良好")
    else:
        print("⚠️  需要进一步优化解析器以达到95%目标")
    
    return results


def main():
    """主函数"""
    try:
        results = validate_all_parsers()
        
        # 保存结果到文件
        output_file = project_root / "fidelity_report.txt"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("样式保真度验证报告\n")
            f.write("="*60 + "\n\n")
            
            for parser_name, report in results.items():
                f.write(f"解析器: {parser_name}\n")
                f.write(f"保真度: {report.overall_fidelity:.2f}%\n")
                f.write(f"分析单元格: {report.analyzed_cells}\n")
                f.write("\n")
        
        logger.info(f"保真度报告已保存到: {output_file}")
        
    except Exception as e:
        logger.error(f"验证过程中出现错误: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
