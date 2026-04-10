import pandas as pd
import os

input_path = r"C:\Users\Administrator\Desktop\工作簿3.xlsx"
output_path = r"C:\Users\Administrator\Desktop\xyy违规行汇总.xlsx"

# 违规词列表（无视大小写）
banned = [
    "MULTICAM", "ALPINE", "ARID", "TROPIC",
    "ATO - ALPINE TERRAIN OPERATIONS", "ATLAS STANDARD",
    "AIRLITE", "HIGHBLEND", "HALFJAK"
]

# banned = [
#     "Mokaloha"
# ]
banned_lower = [w.lower() for w in banned]

def row_has_banned(row):
    # row is a pandas Series
    for cell in row:
        if pd.isna(cell):
            continue
        text = str(cell).lower()
        for term in banned_lower:
            if term in text:
                return True
    return False

def main():
    if not os.path.exists(input_path):
        print(f"未找到输入文件：{input_path}")
        return

    try:
        sheets = pd.read_excel(input_path, sheet_name=None, header=None, engine="openpyxl")
    except Exception as e:
        print("读取 Excel 文件出错:", e)
        return

    collected = []
    total_sheets = len(sheets)
    sheet_idx = 0
    total_bad_rows = 0

    for sheet_name, df in sheets.items():
        sheet_idx += 1
        bad_count = 0
        # 行索引从 0 开始，前 3 行（索引 0,1,2）视为标题，跳过
        nrows = df.shape[0]
        if nrows <= 3:
            # 没有数据行可筛查
            print(f"目前筛查到第{sheet_idx}个表格（{sheet_name}），筛查出违规行有0行（该表无数据行）")
            continue

        for r in range(3, nrows):
            row = df.iloc[r]
            if row_has_banned(row):
                collected.append(list(row.values))
                bad_count += 1

        total_bad_rows += bad_count
        print(f"目前筛查到第{sheet_idx}个表格（{sheet_name}），筛查出违规行有{bad_count}行")

    # 将收集到的行写入新的 Excel（不写 header，不写 index）
    try:
        if collected:
            out_df = pd.DataFrame(collected)
        else:
            # 创建空的 DataFrame（没有行）
            out_df = pd.DataFrame()
        out_df.to_excel(output_path, index=False, header=False, engine="openpyxl")
        print(f"已将所有违规行合并并导出到：{output_path}（共 {total_bad_rows} 行）")
    except Exception as e:
        print("写出结果文件出错:", e)

if __name__ == "__main__":
    main()