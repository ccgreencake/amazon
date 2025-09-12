import pandas as pd
import re

# ========== 配置区域 ==========
input_file = r"C:\Users\Administrator\Downloads\merged_file (10).xlsx"   # 输入文件
output_file = r"C:\Users\Administrator\Downloads\price_stats.xlsx" # 输出文件
# =============================

def clean_days(x):
    """提取上架天数，返回int，无法解析返回None"""
    if pd.isna(x) or x == "-":
        return None
    match = re.search(r"(\d+)", str(x))
    return int(match.group(1)) if match else None

def clean_price(x):
    """提取价格数字"""
    if pd.isna(x):
        return 0.0
    match = re.search(r"([\d\.]+)", str(x))
    return float(match.group(1)) if match else 0.0

def clean_promo(x):
    """提取促销百分比"""
    if pd.isna(x) or not isinstance(x, str):
        return 0.0
    match = re.search(r"(\d+)%", x)
    return float(match.group(1)) / 100 if match else 0.0

def calc_final_price(row):
    """计算最终价格"""
    buybox = clean_price(row["BuyBox价格"])
    promo = clean_promo(row["促销"])
    shipping = row["配送"]

    if pd.isna(shipping):
        shipping_cost = 0.0
    elif shipping.startswith("FBA") or shipping.startswith("FBM(FREE)"):
        shipping_cost = 0.0
    else:
        shipping_cost = clean_price(shipping)

    final_price = buybox * (1 - promo) + shipping_cost
    return round(final_price, 2)

def shipping_type(x):
    """判断配送方式（FBA / FBM）"""
    if pd.isna(x):
        return "其他"
    if str(x).startswith("FBA"):
        return "FBA"
    if str(x).startswith("FBM"):
        return "FBM"
    return "其他"

def price_bins(df):
    """统计价格区间"""
    # 这里你可以自定义区间，比如 0-10, 10-20, 20-30, 30+
    bins = [0, 10, 20, 30, 50, 100, 9999]
    labels = ["0-10", "10-20", "20-30", "30-50", "50-100", "100+"]
    df["价格区间"] = pd.cut(df["价格"], bins=bins, labels=labels, right=False)
    return df

def main():
    # 读取表格
    df = pd.read_excel(input_file)

    # 保留相关列
    keep_cols = ["ASIN", "产品信息（点击可打开亚马逊页面）", "亚马逊URL", "大类排名", "BuyBox价格", "上架时间", "上架天数", "配送", "促销"]
    df = df[keep_cols]

    # 筛选“上架天数 < 30”
    df["上架天数数值"] = df["上架天数"].apply(clean_days)
    df = df[df["上架天数数值"].notna()]
    # df = df[df["上架天数数值"] < 30]

    # 计算价格
    df["价格"] = df.apply(calc_final_price, axis=1)

    # 区分配送类型
    df["配送类型"] = df["配送"].apply(shipping_type)

    # 分价格区间
    df = price_bins(df)

    # 按配送类型 + 价格区间统计
    stats = df.groupby(["配送类型", "价格区间"]).size().reset_index(name="数量")

    # 输出
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="明细", index=False)
        stats.to_excel(writer, sheet_name="价格区间统计", index=False)

    print("处理完成 ✅，结果已保存到", output_file)


if __name__ == "__main__":
    main()
