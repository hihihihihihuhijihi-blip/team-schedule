#!/usr/bin/env python3
"""
招商业务处理工具集 - 参数收集模式

用法:
    python main.py                    # 提示输入参数
    python main.py -m 202601          # 直接指定月份
"""

import argparse
import os
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def print_files_list(month: str):
    """打印需要的文件列表（不检查是否存在，仅提示）"""
    print("=" * 50)
    print(f"  📅 处理月份: {month}")
    print("=" * 50)
    print("\n请提供以下文件:\n")

    files = [
        "原始招商数据（TS{月}-招商业务部.xlsx）",
        "历史清洗数据（上个月的清洗版）",
        "医院信息表",
        "底价表（26年招商业务部底价表.xlsx）",
    ]

    for i, name in enumerate(files, 1):
        print(f"  [{i}] {name}")

    print("\n" + "=" * 50)
    print("  通过参数提供文件:")
    print("  -r 原始数据 -c 历史数据 -H 医院表 -p 底价表")
    print("=" * 50)

    return []


def run_with_files(month: str, files: dict):
    """运行数据清洗脚本"""
    script_path = os.path.join(SCRIPT_DIR, 'ts_data_clean.py')

    cmd = [
        sys.executable, script_path,
        '-m', month,
    ]

    # 添加文件参数
    if files.get('raw_data'):
        cmd.extend(['-r', files['raw_data']])
    if files.get('clean_data'):
        cmd.extend(['-c', files['clean_data']])
    if files.get('hospital'):
        cmd.extend(['-H', files['hospital']])
    if files.get('price'):
        cmd.extend(['-p', files['price']])

    print(f"\n🚀 执行命令: {' '.join(cmd)}\n")
    result = subprocess.run(cmd)
    return result.returncode


def main():
    parser = argparse.ArgumentParser(
        description='招商业务处理工具集',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s -m 202601              # 指定月份，显示文件列表
  %(prog)s -m 202601 --run        # 指定月份并执行（使用默认路径）
        """
    )

    parser.add_argument('-m', '--month', type=str,
                        help='年月，格式为 YYYYMM (如 202601)')

    parser.add_argument('-r', '--raw-data', type=str,
                        help='原始招商数据路径')

    parser.add_argument('-c', '--clean-data', type=str,
                        help='历史清洗数据路径')

    parser.add_argument('-H', '--hospital-table', type=str,
                        help='医院信息表路径')

    parser.add_argument('-p', '--price-table', type=str,
                        help='底价表路径')

    parser.add_argument('--run', action='store_true',
                        help='执行处理（需要提供所有文件参数）')

    args = parser.parse_args()

    # 如果没有月份，提示输入
    if not args.month:
        print("=" * 50)
        print("           招商业务处理工具集")
        print("=" * 50)
        print("\n[1] TS数据清洗 - 合并原始数据与历史清洗数据，匹配医院信息和底价")
        print()
        month_input = input("请输入年月 (YYYYMM, 如 202601): ").strip()
        if not month_input:
            print("❌ 未输入月份，退出")
            sys.exit(1)
        args.month = month_input

    # 验证月份格式
    if not args.month.isdigit() or len(args.month) != 6:
        print(f"❌ 错误: 月份格式不正确，应为 YYYYMM 格式，如 202601")
        sys.exit(1)

    # 显示文件列表
    print_files_list(args.month)

    # 如果指定了 --run，执行处理
    if args.run:
        files = {}
        if args.raw_data:
            files['raw_data'] = args.raw_data
        if args.clean_data:
            files['clean_data'] = args.clean_data
        if args.hospital_table:
            files['hospital'] = args.hospital_table
        if args.price_table:
            files['price'] = args.price_table

        if len(files) == 4:
            return run_with_files(args.month, files)
        else:
            print("\n❌ --run 模式需要提供所有文件参数 (-r, -c, -H, -p)")


if __name__ == "__main__":
    main()
