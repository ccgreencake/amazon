import os
import pandas as pd

# 文件夹路径
FOLDER_PATH = r"F:\Super Browser\Super Browser\youleilei34@outlook.com\DGD广告"

# 输入文件：刚刚合并后的 CSV
INPUT_FILE = "合并结果.csv"
INPUT_PATH = os.path.join(FOLDER_PATH, INPUT_FILE)

# 输出文件
OUTPUT_FILE = "整理结果.csv"
OUTPUT_PATH = os.path.join(FOLDER_PATH, OUTPUT_FILE)


def read_csv_safely(file_path):
    """
    尝试用常见编码读取 CSV
    """
    encodings = ["utf-8-sig", "utf-8", "gbk", "gb18030"]

    for enc in encodings:
        try:
            return pd.read_csv(file_path, dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue

    return pd.read_csv(
        file_path,
        dtype=str,
        encoding="gb18030",
        encoding_errors="ignore"
    )


def to_number(series):
    """
    将字符串数字转成数值
    兼容：
    1,234
    $12.34
    12.34%
    空值
    """
    return pd.to_numeric(
        series.astype(str)
        .str.strip()
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("USD", "", regex=False)
        .str.replace("--", "", regex=False)
        .str.replace("nan", "", regex=False),
        errors="coerce"
    ).fillna(0)


def safe_divide(numerator, denominator):
    """
    安全除法，避免除以 0
    """
    return numerator / denominator.replace(0, pd.NA)


def tidy_csv():
    if not os.path.exists(INPUT_PATH):
        print(f"找不到输入文件：{INPUT_PATH}")
        return

    print(f"正在读取文件：{INPUT_PATH}")

    df = read_csv_safely(INPUT_PATH)

    # 清理列名，去掉前后空格和 BOM
    df.columns = df.columns.astype(str).str.replace("\ufeff", "", regex=False).str.strip()

    required_columns = [
        "顾客搜索词",
        "展示量",
        "点击量",
        "总成本 (USD)",
        "购买量",
        "销售额 (USD)"
    ]

    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        print("表格缺少以下必要列：")
        for col in missing_columns:
            print(col)
        print()
        print("当前表格实际列名如下：")
        for col in df.columns:
            print(col)
        return

    # 去掉空的顾客搜索词
    df["顾客搜索词"] = df["顾客搜索词"].astype(str).str.strip()
    df = df[df["顾客搜索词"] != ""]
    df = df[df["顾客搜索词"].str.lower() != "nan"]

    # 防止重复标题行混进数据
    df = df[df["顾客搜索词"] != "顾客搜索词"]

    # 需要求和的数字列
    sum_columns = [
        "展示量",
        "点击量",
        "总成本 (USD)",
        "购买量",
        "销售额 (USD)"
    ]

    # 转成数字
    for col in sum_columns:
        df[col] = to_number(df[col])

    # 按顾客搜索词分组求和
    grouped = df.groupby("顾客搜索词", as_index=False)[sum_columns].sum()

    # 用归纳后的数字重新计算点击率、CPC、ACOS
    grouped["点击率"] = safe_divide(grouped["点击量"], grouped["展示量"]) * 100
    grouped["CPC (USD)"] = safe_divide(grouped["总成本 (USD)"], grouped["点击量"])
    grouped["ACOS"] = safe_divide(grouped["总成本 (USD)"], grouped["销售额 (USD)"]) * 100

    # 除法结果为空的填 0
    grouped["点击率"] = grouped["点击率"].fillna(0)
    grouped["CPC (USD)"] = grouped["CPC (USD)"].fillna(0)
    grouped["ACOS"] = grouped["ACOS"].fillna(0)

    # 按点击量从多到少排序
    grouped = grouped.sort_values(by="点击量", ascending=False)

    # 整理输出列顺序
    output_df = grouped[
        [
            "顾客搜索词",
            "展示量",
            "点击量",
            "点击率",
            "总成本 (USD)",
            "CPC (USD)",
            "购买量",
            "销售额 (USD)",
            "ACOS"
        ]
    ].copy()

    # 保留两位小数，点击率和 ACOS 加百分号
    output_df["展示量"] = output_df["展示量"].round(2).map(lambda x: f"{x:.2f}")
    output_df["点击量"] = output_df["点击量"].round(2).map(lambda x: f"{x:.2f}")
    output_df["点击率"] = output_df["点击率"].round(2).map(lambda x: f"{x:.2f}%")
    output_df["总成本 (USD)"] = output_df["总成本 (USD)"].round(2).map(lambda x: f"{x:.2f}")
    output_df["CPC (USD)"] = output_df["CPC (USD)"].round(2).map(lambda x: f"{x:.2f}")
    output_df["购买量"] = output_df["购买量"].round(2).map(lambda x: f"{x:.2f}")
    output_df["销售额 (USD)"] = output_df["销售额 (USD)"].round(2).map(lambda x: f"{x:.2f}")
    output_df["ACOS"] = output_df["ACOS"].round(2).map(lambda x: f"{x:.2f}%")

    # 保存
    output_df.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")

    print("整理完成！")
    print(f"整理后的文件已保存到：{OUTPUT_PATH}")


if __name__ == "__main__":
    tidy_csv()