#!/usr/bin/env python3
"""
TS招商数据清洗工具

将原始招商数据与历史清洗数据合并，并匹配医院信息表和底价表。

使用示例:
    python ts_data_clean.py                    # 使用默认参数
    python ts_data_clean.py -m 202601          # 指定月份
    python ts_data_clean.py -m 202601 -r ~/data.xlsx  # 指定输入文件
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
    """获取上个月的年月，格式为 YYYYMM"""
    today = datetime.now()
    # 如果是1月，上个月是去年的12月
    if today.month == 1:
        return f"{today.year - 1}12"
    else:
        return f"{today.year}{today.month - 1:02d}"


def get_default_paths(month: str) -> dict:
    """根据月份生成默认文件路径"""
    desktop = Path.home() / "Desktop"

    # 解析年月
    year = month[:4]
    month_num = month[4:6]

    # 计算上个月（用于历史清洗数据路径）
    month_int = int(month_num)
    if month_int == 1:
        prev_year = str(int(year) - 1)
        prev_month = "12"
    else:
        prev_year = year
        prev_month = f"{month_int - 1:02d}"

    # 计算医院表月份（处理月份+1）
    hospital_month = month_int + 1
    hospital_year = int(year[-2:])
    if hospital_month > 12:
        hospital_month = 1
        hospital_year += 1

    return {
        "raw_data": desktop / f"TS{month}-招商业务部.xlsx",
        "clean_data": desktop / "Commercial Excellence 2025" / "BI 基础表" / "2025招商业务部" / f"TS{prev_year}{prev_month}-招商（清洗版）.xlsx",
        "hospital": desktop / "Commercial Excellence 2025" / "销售数据管理" / "医院信息表" / f"{hospital_year}-{hospital_month}医院表.xlsx",
        "price": desktop / "26年招商业务部底价表.xlsx",
        "output_sync": desktop / f"TS{month}-招商.xlsx",
        "output_final": desktop / f"TS{month}-招商（已清洗）.xlsx",
    }


def check_file_exists(path: Path, description: str) -> bool:
    """检查文件是否存在，不存在则提示"""
    if not path.exists():
        print(f"❌ 错误: {description} 文件不存在: {path}")
        return False
    return True


def sync_data(raw_data_path: Path, clean_data_path: Path, month: int, output_path: Path) -> pd.DataFrame:
    """
    步骤1: 同步数据
    - 读取原始招商数据文件
    - 筛选指定年月的数据
    - 追加到已清洗的历史数据中
    """
    print(f"\n📖 读取原始数据: {raw_data_path}")
    raw_df = pd.read_excel(raw_data_path, engine='openpyxl')

    print(f"📖 读取历史清洗数据: {clean_data_path}")
    clean_df = pd.read_excel(clean_data_path, engine='openpyxl')

    # 手动对齐列名
    raw_df.rename(columns={
        "年月": "年月",
        "医院（备案医院）": "原医院（备案医院）",
        "省区": "原省区",
        "片区": "原片区",
        "皮肤-终端类型": "皮肤-终端类型以此为准！"
    }, inplace=True)

    # 筛选指定年月的数据
    filtered_raw_df = raw_df[raw_df['年月'] == month]

    print(f"✅ 筛选出 {len(filtered_raw_df)} 条年月为 {month} 的数据")

    # 合并数据
    updated_df = pd.concat([clean_df, filtered_raw_df], ignore_index=True)

    print(f"✅ 合并后共 {len(updated_df)} 条数据")

    # 将时间相关列转换为文本类型
    text_cols = ['年月', '月份', '双月', '季度', '年度', '月度', '双月份', '季度份']
    for col in text_cols:
        if col in updated_df.columns:
            updated_df[col] = updated_df[col].astype(str)

    # 导出中间结果
    print(f"💾 导出同步后的数据到: {output_path}")
    updated_df.to_excel(output_path, index=False)
    # 应用 Excel 格式设置
    apply_excel_formatting(output_path)

    return updated_df


def match_data(
    df: pd.DataFrame,
    hospital_path: Path,
    price_path: Path,
    month: int,
    output_path: Path
) -> pd.DataFrame:
    """
    步骤2: 匹配数据
    - 匹配医院信息表（CRM医院编码）
    - 匹配底价表（26年底价）
    - 计算实际底价和底价额
    """
    print(f"\n📖 读取医院信息表: {hospital_path}")
    hospital_df = pd.read_excel(hospital_path, sheet_name='医院管理', engine='openpyxl')

    print(f"📖 读取底价表: {price_path}")
    price_df = pd.read_excel(price_path, sheet_name='底价匹配', engine='openpyxl')

    # 去除列名中的空格和换行符
    df.columns = df.columns.str.strip()
    hospital_df.columns = hospital_df.columns.str.strip()
    price_df.columns = price_df.columns.str.strip()

    # ==========================================
    # 基础字段更新 (针对指定月份)
    # ==========================================
    # 注意：年月列在sync_data中已转为字符串，需要用字符串比较
    mask = df['年月'].astype(str) == str(month)

    df.loc[mask, '片区'] = df.loc[mask, '原片区']
    df.loc[mask, '省区'] = df.loc[mask, '原省区']
    df.loc[mask, 'CRM医院名称'] = df.loc[mask, '原医院（备案医院）']
    df.loc[mask, '代理商(剔除票折)'] = df.loc[mask, '代理商']
    df.loc[mask, '类型'] = df.loc[mask, '皮肤-终端类型以此为准！']

    # 从医院表中提取医院名称和CRM医院编码的对应关系
    # 兼容不同列名：优先使用CRM医院名称，否则使用医院名称
    if 'CRM医院名称' in hospital_df.columns:
        hospital_code_df = hospital_df[['CRM医院名称', 'CRM医院编码']].dropna()
        name_col = 'CRM医院名称'
    elif '医院名称' in hospital_df.columns:
        hospital_code_df = hospital_df[['医院名称', 'CRM医院编码']].dropna()
        name_col = '医院名称'
    else:
        raise ValueError("医院表中缺少必要的列（CRM医院名称 或 医院名称）")
    code_mapping = dict(zip(hospital_code_df[name_col], hospital_code_df['CRM医院编码']))

    # 匹配CRM医院编码
    df.loc[mask, 'CRM医院编码'] = df.loc[mask, 'CRM医院名称'].map(code_mapping)
    print(f"✅ 匹配 CRM医院编码")

    # 统一类型名称
    df['皮肤-终端类型以此为准！'] = df['皮肤-终端类型以此为准！'].replace({'医疗端': '医疗', '非医疗端': '非医疗'})
    df['类型'] = df['类型'].replace({'医疗端': '医疗', '非医疗端': '非医疗'})

    # ==========================================
    # 底价匹配逻辑 (不再区分医疗/非医疗)
    # ==========================================
    # 准备底价字典 - 只取'商名S'和'底价价格'
    price_map_df = price_df[['商名S', '底价价格']].drop_duplicates(subset=['商名S'])
    price_dict = dict(zip(price_map_df['商名S'], price_map_df['底价价格']))

    # 匹配26年底价（直接用'商名S'进行map，不再拼接类型）
    df.loc[mask, '26年底价'] = df.loc[mask, '商名S'].map(price_dict)
    print(f"✅ 匹配 26年底价")

    # 计算26年实际底价（比较单价和26年底价，取较小值）
    df.loc[mask, '26年实际底价'] = df.loc[mask].apply(
        lambda row: row['单价'] if row['单价'] < row['26年底价'] else row['26年底价'],
        axis=1
    )

    # 数据清洗其他列
    df.loc[mask, '数据口径'] = '保留'
    df.loc[mask, '代理商(剔除票折)'] = df.loc[mask, '代理商(剔除票折)'].str.replace('（票折）', '')

    # ==========================================
    # 特殊底价修正区域 (优先级最高，覆盖前面的计算)
    # ==========================================
    # 吡美莫司14g - 单价是 39.34 或 42 的保持原单价
    condition_pime_1 = (df['商名S'] == '吡美莫司14g') & (df['单价'].isin([39.34, 42])) & mask
    df.loc[condition_pime_1, '26年实际底价'] = df.loc[condition_pime_1, '单价']

    # 吡美莫司14g + 上海（单价不是 39.34 或 42）
    condition_pime_2 = (df['商名S'] == '吡美莫司14g') & (df['省份'] == '上海') & (~df['单价'].isin([39.34, 42])) & mask
    df.loc[condition_pime_2, '26年实际底价'] = 38

    # 江西 他克莫司10g
    condition_4 = (df['商名S'] == '他克莫司10g') & (df['省份'] == '江西') & mask
    df.loc[condition_4, '26年实际底价'] = 15.6

    # 江西 他克莫司10g/0.03%
    condition_5 = (df['商名S'] == '他克莫司10g/0.03%') & (df['省份'] == '江西') & mask
    df.loc[condition_5, '26年实际底价'] = 9.48

    print(f"✅ 应用特殊产品底价调整")

    # ==========================================
    # 金额计算区域 (确保在所有底价修正后执行)
    # ==========================================
    # 计算26年底价额
    df.loc[mask, '26年底价额'] = df.loc[mask, '26年实际底价'] * df.loc[mask, '进货数量']

    # 计算26年底价额(万)
    df.loc[mask, '26年底价额(万)'] = df.loc[mask, '26年底价额'] / 10000

    # 将时间相关列转换为文本类型
    text_cols = ['年月', '月份', '双月', '季度', '年度', '月度', '双月份', '季度份']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str)

    # 导出最终结果
    print(f"\n💾 导出清洗后的数据到: {output_path}")
    df.to_excel(output_path, index=False)
    # 应用 Excel 格式设置
    apply_excel_formatting(output_path)

    return df


def main():
    parser = argparse.ArgumentParser(
        description='TS招商数据清洗工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s                              # 使用默认参数（上个月）
  %(prog)s -m 202601                    # 指定月份
  %(prog)s -m 202601 -r ~/data.xlsx     # 指定输入文件
  %(prog)s -m 202601 -o output.xlsx     # 指定输出文件名
        """
    )

    parser.add_argument('-m', '--month', type=str,
                        default=get_default_month(),
                        help='年月，格式为 YYYYMM（默认为上个月）')

    parser.add_argument('-r', '--raw-data', type=str,
                        help='原始招商数据路径')

    parser.add_argument('-c', '--clean-data', type=str,
                        help='已清洗历史数据路径')

    parser.add_argument('-H', '--hospital-table', type=str,
                        help='医院信息表路径')

    parser.add_argument('-p', '--price-table', type=str,
                        help='底价表路径')

    parser.add_argument('-o', '--output', type=str,
                        help='输出文件名（不含扩展名，默认生成两个文件）')

    parser.add_argument('-d', '--output-dir', type=str,
                        help='输出目录')

    parser.add_argument('--dry-run', action='store_true',
                        help='只验证文件路径，不执行处理')

    args = parser.parse_args()

    # 验证月份格式
    if not args.month.isdigit() or len(args.month) != 6:
        print(f"❌ 错误: 月份格式不正确，应为 YYYYMM 格式，如 202601")
        sys.exit(1)

    month_int = int(args.month)

    # 检查必须提供文件参数（不使用默认路径）
    missing = []
    if not args.raw_data:
        missing.append("--raw-data (-r)")
    if not args.clean_data:
        missing.append("--clean-data (-c)")
    if not args.hospital_table:
        missing.append("--hospital-table (-H)")
    if not args.price_table:
        missing.append("--price-table (-p)")

    if missing:
        print("❌ 错误: 缺少必须的文件参数")
        print("\n必须提供以下文件:")
        for m in missing:
            print(f"  {m}")
        print("\n示例:")
        print(f"  python ts_data_clean.py -m {args.month} \\")
        print(f"    -r 原始数据.xlsx \\")
        print(f"    -c 历史数据.xlsx \\")
        print(f"    -H 医院表.xlsx \\")
        print(f"    -p 底价表.xlsx")
        sys.exit(1)

    # 确定实际使用的路径
    raw_data_path = Path(args.raw_data)
    clean_data_path = Path(args.clean_data)
    hospital_path = Path(args.hospital_table)
    price_path = Path(args.price_table)

    # 确定输出路径
    if args.output_dir:
        output_dir = Path(args.output_dir)
    elif args.raw_data:
        # 默认输出到输入文件所在目录
        output_dir = Path(args.raw_data).parent
    else:
        output_dir = Path.home() / "Desktop"

    output_name = args.output if args.output else f"TS{args.month}-招商"
    output_sync_path = output_dir / f"{output_name}.xlsx"
    output_final_path = output_dir / f"{output_name}（已清洗）.xlsx"

    # 显示配置
    print("=" * 50)
    print(f"📅 处理月份: {args.month}")
    print("=" * 50)
    print(f"原始招商数据: {raw_data_path}")
    print(f"历史清洗数据: {clean_data_path}")
    print(f"医院信息表:   {hospital_path}")
    print(f"底价表:       {price_path}")
    print(f"输出文件:     {output_final_path}")
    print("=" * 50)

    # 检查文件存在性
    all_exist = True
    all_exist &= check_file_exists(raw_data_path, "原始招商数据")
    all_exist &= check_file_exists(clean_data_path, "历史清洗数据")
    all_exist &= check_file_exists(hospital_path, "医院信息表")
    all_exist &= check_file_exists(price_path, "底价表")

    if not all_exist:
        sys.exit(1)

    if args.dry_run:
        print("\n✅ 文件检查通过，未执行处理（--dry-run 模式）")
        sys.exit(0)

    # 步骤1: 同步数据
    df = sync_data(raw_data_path, clean_data_path, month_int, output_sync_path)

    # 步骤2: 匹配数据
    df = match_data(df, hospital_path, price_path, month_int, output_final_path)

    # 完成
    print("=" * 50)
    print(f"✅ 处理完成！共 {len(df)} 条数据")
    print(f"📊 其中年月 {month_int} 的数据: {len(df[df['年月'] == str(month_int)])} 条")
    print("=" * 50)


if __name__ == "__main__":
    main()
