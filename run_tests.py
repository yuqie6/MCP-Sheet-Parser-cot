#!/usr/bin/env python3
"""
测试运行脚本

使用uv包管理器运行项目测试，支持不同的测试模式和覆盖率报告。
"""

import subprocess
import sys
import argparse
from pathlib import Path


def run_command(cmd: list[str], description: str) -> bool:
    """运行命令并处理结果。"""
    print(f"\n🔄 {description}")
    print(f"执行命令: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(f"✅ {description} 成功完成")
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} 失败")
        print(f"错误代码: {e.returncode}")
        if e.stdout:
            print("标准输出:")
            print(e.stdout)
        if e.stderr:
            print("错误输出:")
            print(e.stderr)
        return False


def main():
    """主函数。"""
    parser = argparse.ArgumentParser(description="运行项目测试")
    parser.add_argument(
        "--mode", 
        choices=["core", "all", "coverage"], 
        default="core",
        help="测试模式: core=只测试核心服务, all=所有测试, coverage=生成覆盖率报告"
    )
    parser.add_argument(
        "--verbose", "-v", 
        action="store_true",
        help="详细输出"
    )
    parser.add_argument(
        "--target-coverage", 
        type=int, 
        default=80,
        help="目标覆盖率百分比 (默认: 80)"
    )
    
    args = parser.parse_args()
    
    # 确保在项目根目录
    project_root = Path(__file__).parent
    if not (project_root / "pyproject.toml").exists():
        print("❌ 错误: 请在项目根目录运行此脚本")
        sys.exit(1)
    
    print("🧪 MCP Sheet Parser 测试运行器")
    print(f"📁 项目目录: {project_root}")
    print(f"🎯 目标覆盖率: {args.target_coverage}%")
    
    success = True
    
    if args.mode == "core":
        # 只测试核心服务
        cmd = [
            "uv", "run", "pytest", 
            "tests/test_core_service.py",
            "--cov=src.core_service",
            "--cov-report=term-missing"
        ]
        if args.verbose:
            cmd.append("-v")
        
        success = run_command(cmd, "运行核心服务测试")
        
    elif args.mode == "all":
        # 运行所有测试
        cmd = [
            "uv", "run", "pytest",
            "--cov=src",
            "--cov-report=term-missing"
        ]
        if args.verbose:
            cmd.append("-v")
        
        success = run_command(cmd, "运行所有测试")
        
    elif args.mode == "coverage":
        # 生成详细的覆盖率报告
        cmd = [
            "uv", "run", "pytest",
            "--cov=src",
            "--cov-report=term-missing",
            "--cov-report=html",
            f"--cov-fail-under={args.target_coverage}"
        ]
        if args.verbose:
            cmd.append("-v")
        
        success = run_command(cmd, f"生成覆盖率报告 (目标: {args.target_coverage}%)")
        
        if success:
            html_report = project_root / "htmlcov" / "index.html"
            if html_report.exists():
                print(f"📊 HTML覆盖率报告已生成: {html_report}")
                print("💡 提示: 在浏览器中打开此文件查看详细覆盖率报告")
    
    # 输出结果
    if success:
        print("\n🎉 测试运行成功!")
        if args.mode == "core":
            print("✅ 核心服务测试通过，覆盖率达到目标")
        elif args.mode == "all":
            print("✅ 所有测试通过")
        elif args.mode == "coverage":
            print(f"✅ 覆盖率达到 {args.target_coverage}% 目标")
    else:
        print("\n💥 测试运行失败!")
        sys.exit(1)


if __name__ == "__main__":
    main()
