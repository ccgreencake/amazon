import pandas as pd
import os
import re

# 文件夹路径
folder_path = r"C:\Users\Administrator\Desktop\order"

for file_name in os.listdir(folder_path):
    if file_name.lower().endswith(".txt"):
        txt_path = os.path.join(folder_path, file_name)
        xlsx_path = os.path.splitext(txt_path)[0] + ".xlsx"

        rows = []

        with open(txt_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                # 自动按 制表符 / 逗号 / 多个空格 分割
                row = re.split(r"\t|,|\s{2,}", line)
                rows.append(row)

        if rows:
            df = pd.DataFrame(rows)
            df.to_excel(xlsx_path, index=False, header=False)
            print(f"已转换：{file_name}")
        else:
            print(f"跳过空文件：{file_name}")

print("全部 TXT 文件转换完成")
