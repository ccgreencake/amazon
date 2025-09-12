import pandas as pd
import re

# ========== 配置区域 ==========
input_file = r"C:\Users\Administrator\Downloads\merged_file (16).xlsx"   # 输入文件
output_file = r"C:\Users\Administrator\Downloads\output0912 .xlsx" # 输出文件
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
    elif shipping in ["FBA", "FBM(FREE)"]:
        shipping_cost = 0.0
    else:
        shipping_cost = clean_price(shipping)

    final_price = buybox * (1 - promo) + shipping_cost
    return round(final_price, 2)

def word_frequency(series):
    """统计词频（单词 & 双词），并排序"""
    from collections import Counter

    all_words = []
    two_words = []

    for text in series.dropna():
        # 只保留字母数字，分词
        words = re.findall(r"[a-zA-Z0-9]+", str(text).lower())
        all_words.extend(words)
        two_words.extend([" ".join(words[i:i+2]) for i in range(len(words)-1)])

    one_word_count = Counter(all_words)
    two_word_count = Counter(two_words)

    # 转成 DataFrame 并排序
    df_one = pd.DataFrame(one_word_count.items(), columns=["单个词", "词频"]).sort_values(by="词频", ascending=False)
    df_two = pd.DataFrame(two_word_count.items(), columns=["两个词", "两个词词频"]).sort_values(by="两个词词频", ascending=False)

    # 合并成一个表（左右拼接）
    result = pd.concat([df_one.reset_index(drop=True), df_two.reset_index(drop=True)], axis=1)
    return result

def main():
    # 读取表格
    df = pd.read_excel(input_file)

    # 筛选列
    keep_cols = ["ASIN", "产品信息（点击可打开亚马逊页面）", "亚马逊URL", "大类排名", "BuyBox价格", "上架时间", "上架天数", "配送", "促销"]
    df = df[keep_cols]

    # 筛选“上架天数 < 30”的行，去掉"-"
    df["上架天数数值"] = df["上架天数"].apply(clean_days)
    df = df[df["上架天数数值"].notna()]
    df = df[df["上架天数数值"] < 30]

    # 计算价格
    df["价格"] = df.apply(calc_final_price, axis=1)

    # 排序：先按天数，再按大类排名
    df["大类排名数值"] = df["大类排名"].apply(lambda x: clean_days(x))  # 提取大类排名的数字
    df = df.sort_values(by=["上架天数数值", "大类排名数值"], ascending=[True, True])

    # 词频统计子表（已排序）
    sub_df = word_frequency(df["产品信息（点击可打开亚马逊页面）"])

    # 输出到Excel（多sheet）
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="主表", index=False)
        sub_df.to_excel(writer, sheet_name="词频统计", index=False)

    print("处理完成 ✅，结果已保存到", output_file)


if __name__ == "__main__":
    main()
