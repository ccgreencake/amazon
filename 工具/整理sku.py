import os
import re
import pandas as pd

# 原始文件路径
INPUT_PATH = r"C:\Users\Administrator\Downloads\product_20260604113655418.xls"

# 输出文件路径
OUTPUT_PATH = r"C:\Users\Administrator\Downloads\product_整理结果.xlsx"


def clean_sheet_name(name, used_names):
    """
    清理 Excel 子表名称
    """
    if pd.isna(name) or str(name).strip() == "":
        name = "未分类"
    else:
        name = str(name).strip()

    name = re.sub(r'[\\/*?:\[\]]', "_", name)
    name = name[:31]

    if name == "":
        name = "未分类"

    original_name = name
    count = 1
    while name in used_names:
        count += 1
        suffix = f"_{count}"
        name = original_name[:31 - len(suffix)] + suffix

    used_names.add(name)
    return name


def apply_excel_format(writer, sheet_name, df):
    """
    按要求设置 Excel 格式：
    1. 除标题行外，所有行高 80 磅
    2. SKU列宽 80 磅、产品缩略图列宽 80 磅
    3. 浏览量列宽 50 磅、是否禁售列宽 50 磅
    """
    worksheet = writer.sheets[sheet_name]

    # --- 1. 设置行高 ---
    # 跳过第1行标题，所有数据行高设置为80磅
    for row_num in range(2, len(df) + 2):
        worksheet.row_dimensions[row_num].height = 80

    # --- 2. 设置列宽 ---
    # 定义各列固定宽度配置
    width_config = {
        "SKU": 80,
        "产品缩略图": 80,
        "浏览量": 50,
        "是否禁售": 50
    }

    # 遍历列，应用对应宽度
    for col_idx, col_name in enumerate(df.columns, 1):
        col_letter = worksheet.cell(row=1, column=col_idx).column_letter
        target_width = width_config.get(col_name, 30)  # 无配置列默认30磅
        worksheet.column_dimensions[col_letter].width = target_width


def read_excel_safely(file_path):
    """
    自动识别格式读取 Excel
    """
    try:
        # 针对 xlsx 格式（哪怕后缀写的是 .xls）
        return pd.read_excel(file_path, dtype=str, engine="openpyxl")
    except Exception:
        # 针对真正的旧版 xls 格式
        return pd.read_excel(file_path, dtype=str, engine="xlrd")


def main():
    if not os.path.exists(INPUT_PATH):
        print(f"找不到文件：{INPUT_PATH}")
        return

    print("正在读取表格数据...")

    try:
        df = read_excel_safely(INPUT_PATH)
    except Exception as e:
        print(f"读取失败：{e}")
        return

    # 清理列名（防止有空格或特殊字符）
    df.columns = df.columns.astype(str).str.strip()

    required_columns = ["产品父类", "SKU", "产品缩略图", "浏览量", "是否禁售"]

    # 检查列是否存在
    for col in required_columns:
        if col not in df.columns:
            print(f"错误：表格中找不到列 '{col}'")
            print(f"当前表格列名有：{list(df.columns)}")
            return

    # 预处理：填充空值，统一转为字符串
    df["产品父类"] = df["产品父类"].fillna("未分类").astype(str).str.strip()
    df.loc[df["产品父类"] == "", "产品父类"] = "未分类"

    output_columns = ["SKU", "产品缩略图", "浏览量", "是否禁售"]
    used_sheet_names = set()

    print("开始导出并设置格式...")

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        for product_parent, group_df in df.groupby("产品父类", sort=False):
            sheet_name = clean_sheet_name(product_parent, used_sheet_names)

            # 只取需要的列
            final_df = group_df[output_columns].copy()

            # 写入数据
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # 调用格式设置函数
            apply_excel_format(writer, sheet_name, final_df)

            print(f"已完成子表：{sheet_name} (共 {len(final_df)} 行)")

    print("\n" + "=" * 30)
    print("全部整理完成！")
    print(f"输出路径：{OUTPUT_PATH}")
    print("=" * 30)


if __name__ == "__main__":
    main()