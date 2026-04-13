#!/usr/bin/env python3
"""
底价销售进度工具

统计26年招商业务部各省区底价销售达成情况。

使用示例:
    python price_sales_progress.py -m 202601 \\
        -t 26年招商销售进度公示.xlsx \\
        -f TS202601-招商（已清洗）.xlsx \\
        -i T2026-招商-一维表.xlsx
"""

import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# 导入 Excel 格式设置模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel_formatter import apply_excel_formatting


def get_default_month():
    """获取当前年月，格式为 YYYYMM"""
    today = datetime.now()
    return f"{today.year}{today.month:02d}"


def check_file_exists(path: Path, description: str) -> bool:
    """检查文件是否存在，不存在则提示"""
    if not path.exists():
        print(f"❌ 错误: {description} 文件不存在: {path}")
        return False
    return True


def calculate_previous_year_month(month: str) -> str:
    """计算去年同月的年月"""
    year = int(month[:4])
    month_num = int(month[4:6])
    return f"{year - 1}{month_num:02d}"


def calculate_month_range(start_month: str, end_month: str) -> list:
    """
    计算从起始月到结束月的所有年月列表
    格式: YYYYMM
    """
    start_year = int(start_month[:4])
    start_m = int(start_month[4:6])
    end_year = int(end_month[:4])
    end_m = int(end_month[4:6])

    months = []
    for year in range(start_year, end_year + 1):
        m_start = start_m if year == start_year else 1
        m_end = end_m if year == end_year else 12
        for m in range(m_start, m_end + 1):
            months.append(f"{year}{m:02d}")
    return months


def load_template(template_path: Path) -> pd.DataFrame:
    """
    读取主表（26年招商销售进度公示.xlsx）
    包含：片区、招商经理、26年部门考核指标
    """
    print(f"\n📖 读取主表: {template_path}")
    df = pd.read_excel(template_path)
    print(f"✅ 主表形状: {df.shape}")
    return df


def load_indicator_table(indicator_path: Path, current_month: str) -> dict:
    """
    读取指标表（T2026-招商-一维表.xlsx）
    计算各省区从202601到当前月的累计指标（仅医疗端）

    返回: {省区: 累计指标值}
    """
    print(f"\n📖 读取指标表: {indicator_path}")
    df = pd.read_excel(indicator_path)

    # 去除列名中的空格
    df.columns = df.columns.str.strip()

    print(f"✅ 指标表形状: {df.shape}")

    # 计算需要累加的月份范围
    months_in_range = calculate_month_range("202601", current_month)
    print(f"📅 累计指标月份范围: {months_in_range}")

    # 筛选指定月份范围的数据
    df_filtered = df[df['年月'].astype(str).isin(months_in_range)]

    # 筛选医疗端数据
    df_filtered = df_filtered[df_filtered['渠道类型'] == '医疗']

    # 按省区汇总底价指标（作为累计指标）
    indicator_by_region = df_filtered.groupby('省区')['底价指标'].sum().to_dict()

    print(f"✅ 各省区累计指标计算完成（医疗端）")
    for region, value in sorted(indicator_by_region.items()):
        if pd.notna(value) and value > 0:
            print(f"   {region}: {value:.2f}")

    return indicator_by_region


def load_flow_data(flow_path: Path) -> pd.DataFrame:
    """
    读取流向数据（TS{月份}-招商（已清洗）.xlsx）
    """
    print(f"\n📖 读取流向数据: {flow_path}")
    df = pd.read_excel(flow_path)
    df.columns = df.columns.str.strip()
    print(f"✅ 流向数据形状: {df.shape}")
    return df


def calculate_cumulative_sales(
    df: pd.DataFrame,
    start_month: str,
    end_month: str,
    filter_medical: bool = True
) -> dict:
    """
    计算累计销售额

    参数:
        df: 流向数据
        start_month: 起始年月 (YYYYMM)
        end_month: 结束年月 (YYYYMM)
        filter_medical: 是否筛选医疗终端和保留数据

    返回: {省区: 累计销售额}
    """
    months_in_range = calculate_month_range(start_month, end_month)

    # 筛选年月范围
    df_filtered = df[df['年月'].astype(str).isin(months_in_range)]

    # 筛选医疗终端和保留数据
    if filter_medical:
        df_filtered = df_filtered[
            (df_filtered['皮肤-终端类型以此为准！'] == '医疗') &
            (df_filtered['数据口径'] == '保留')
        ]

    # 按省区汇总 26年底价额(万)
    sales_by_region = df_filtered.groupby('省区')['26年底价额(万)'].sum().to_dict()

    return sales_by_region


def build_output_df(
    template_df: pd.DataFrame,
    indicator_by_region: dict,
    sales_2026_by_region: dict,
    sales_2025_by_region: dict,
    current_month: str
) -> pd.DataFrame:
    """
    构建输出数据框

    计算字段:
    - 26年累计指标
    - 26年累计销售额
    - 25年累计销售额
    - 26年累计达成率 = 26年累计销售额 / 26年累计指标
    - 26年销售进度 = 26年累计销售额 / 26年部门考核指标
    - 26年累计同比增长额 = 26年累计销售额 - 25年累计销售额
    """
    # 按照主表的省区顺序构建输出
    result_rows = []

    for _, row in template_df.iterrows():
        region = row['省区']  # 主表中的"省区"

        # 跳过主表中的总计行（会重新添加）
        if region == '总计' or pd.isna(region):
            continue

        dept_target = row['26年部门考核指标']

        # 获取各项数据
        cumulative_indicator = indicator_by_region.get(region, 0)
        sales_2026 = sales_2026_by_region.get(region, 0)
        sales_2025 = sales_2025_by_region.get(region, 0)

        # 计算派生指标
        achievement_rate = (sales_2026 / cumulative_indicator * 100) if cumulative_indicator > 0 else 0
        sales_progress = (sales_2026 / dept_target * 100) if dept_target > 0 else 0
        yoy_growth = sales_2026 - sales_2025

        result_rows.append({
            '省区': region,
            '招商经理': row['招商经理'],
            '26年部门考核指标': round(dept_target, 1),
            '26年累计指标': round(cumulative_indicator, 1),
            '25年累计销售额': round(sales_2025, 1),
            '26年累计销售额': round(sales_2026, 1),
            '26年累计达成率': f"{achievement_rate:.2f}%",
            '26年销售进度': f"{sales_progress:.2f}%",
            '26年累计同比增长额': round(yoy_growth, 1),
        })

    result_df = pd.DataFrame(result_rows)

    # 按26年销售进度倒序排列（总计行最后添加，不参与排序）
    # 将百分比字符串转换为数值进行排序
    result_df['_sort_key'] = result_df['26年销售进度'].str.replace('%', '').astype(float)
    result_df = result_df.sort_values('_sort_key', ascending=False)
    result_df = result_df.drop('_sort_key', axis=1)
    result_df = result_df.reset_index(drop=True)

    # 添加总计行 - 使用原始数据计算，避免四舍五入误差
    # 从主表直接求和（排除总计行）
    total_dept_target = template_df[template_df['省区'] != '总计']['26年部门考核指标'].sum()
    # 从原始字典直接求和（未四舍五入的数据）
    total_indicator = sum(indicator_by_region.values())
    total_sales_2026 = sum(sales_2026_by_region.values())
    total_sales_2025 = sum(sales_2025_by_region.values())

    total_achievement_rate = (total_sales_2026 / total_indicator * 100) if total_indicator > 0 else 0
    total_sales_progress = (total_sales_2026 / total_dept_target * 100) if total_dept_target > 0 else 0

    total_row = {
        '省区': '总计',
        '招商经理': '',
        '26年部门考核指标': round(total_dept_target, 1),
        '26年累计指标': round(total_indicator, 1),
        '25年累计销售额': round(total_sales_2025, 1),
        '26年累计销售额': round(total_sales_2026, 1),
        '26年累计达成率': f"{total_achievement_rate:.2f}%",
        '26年销售进度': f"{total_sales_progress:.2f}%",
        '26年累计同比增长额': round(total_sales_2026 - total_sales_2025, 1),
    }

    result_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)

    return result_df


def main():
    parser = argparse.ArgumentParser(
        description='底价销售进度工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s -m 202601 \\
    -t 26年招商销售进度公示.xlsx \\
    -f TS202601-招商（已清洗）.xlsx \\
    -i T2026-招商-一维表.xlsx
        """
    )

    parser.add_argument('-m', '--month', type=str,
                        default=get_default_month(),
                        help='当前年月，格式为 YYYYMM（默认为当前月）')

    parser.add_argument('-t', '--template', type=str,
                        help='主表路径（26年招商销售进度公示.xlsx）')

    parser.add_argument('-f', '--flow-data', type=str,
                        help='流向数据路径（TS{月份}-招商（已清洗）.xlsx）')

    parser.add_argument('-i', '--indicator-table', type=str,
                        help='指标表路径（T2026-招商-一维表.xlsx）')

    parser.add_argument('-o', '--output', type=str,
                        help='输出文件名（不含扩展名）')

    parser.add_argument('-d', '--output-dir', type=str,
                        help='输出目录')

    args = parser.parse_args()

    # 验证月份格式
    if not args.month.isdigit() or len(args.month) != 6:
        print(f"❌ 错误: 月份格式不正确，应为 YYYYMM 格式，如 202601")
        sys.exit(1)

    current_month = args.month
    previous_year_month = calculate_previous_year_month(current_month)

    # 检查必须提供文件参数
    missing = []
    if not args.template:
        missing.append("--template (-t)")
    if not args.flow_data:
        missing.append("--flow-data (-f)")
    if not args.indicator_table:
        missing.append("--indicator-table (-i)")

    if missing:
        print("❌ 错误: 缺少必须的文件参数")
        print("\n必须提供以下文件:")
        for m in missing:
            print(f"  {m}")
        print("\n示例:")
        print(f"  python price_sales_progress.py -m {current_month} \\")
        print(f"    -t 26年招商销售进度公示.xlsx \\")
        print(f"    -f TS{current_month}-招商（已清洗）.xlsx \\")
        print(f"    -i T2026-招商-一维表.xlsx")
        sys.exit(1)

    # 确定实际使用的路径
    template_path = Path(args.template)
    flow_data_path = Path(args.flow_data)
    indicator_table_path = Path(args.indicator_table)

    # 确定输出路径
    if args.output_dir:
        output_dir = Path(args.output_dir)
    elif args.flow_data:
        output_dir = Path(args.flow_data).parent
    else:
        output_dir = Path.home() / "Desktop"

    output_name = args.output if args.output else f"26年招商销售进度公示_{current_month}"
    output_path = output_dir / f"{output_name}.xlsx"

    # 显示配置
    print("=" * 60)
    print(f"📅 统计月份: {current_month}")
    print("=" * 60)
    print(f"主表:         {template_path}")
    print(f"流向数据:     {flow_data_path}")
    print(f"指标表:       {indicator_table_path}")
    print(f"输出文件:     {output_path}")
    print("=" * 60)

    # 检查文件存在性
    all_exist = True
    all_exist &= check_file_exists(template_path, "主表")
    all_exist &= check_file_exists(flow_data_path, "流向数据")
    all_exist &= check_file_exists(indicator_table_path, "指标表")

    if not all_exist:
        sys.exit(1)

    # 步骤1: 读取主表
    template_df = load_template(template_path)

    # 步骤2: 读取指标表并计算累计指标
    indicator_by_region = load_indicator_table(indicator_table_path, current_month)

    # 步骤3: 读取流向数据
    flow_df = load_flow_data(flow_data_path)

    # 步骤4: 计算26年累计销售额（从202601到当前月）
    print(f"\n📊 计算26年累计销售额 (202601 ~ {current_month})")
    sales_2026_by_region = calculate_cumulative_sales(
        flow_df, "202601", current_month, filter_medical=True
    )
    for region, value in sorted(sales_2026_by_region.items()):
        if pd.notna(value) and value > 0:
            print(f"   {region}: {value:.2f} 万")

    # 步骤5: 计算25年同期累计销售额
    prev_year_start = calculate_previous_year_month("202601")
    prev_year_end = previous_year_month
    print(f"\n📊 计算25年同期累计销售额 ({prev_year_start} ~ {prev_year_end})")
    sales_2025_by_region = calculate_cumulative_sales(
        flow_df, prev_year_start, prev_year_end, filter_medical=True
    )
    for region, value in sorted(sales_2025_by_region.items()):
        if pd.notna(value) and value > 0:
            print(f"   {region}: {value:.2f} 万")

    # 步骤6: 构建输出数据框
    print(f"\n🔨 构建输出数据")
    result_df = build_output_df(
        template_df,
        indicator_by_region,
        sales_2026_by_region,
        sales_2025_by_region,
        current_month
    )

    # 导出结果
    print(f"\n💾 导出结果到: {output_path}")
    result_df.to_excel(output_path, index=False)
    # 应用 Excel 格式设置
    apply_excel_formatting(output_path)

    # 完成
    print("=" * 60)
    print(f"✅ 处理完成！共 {len(result_df)} 行数据")
    print("=" * 60)
    print("\n结果预览:")
    print(result_df.to_string(index=False))


if __name__ == "__main__":
    main()
