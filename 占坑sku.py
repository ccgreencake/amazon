import os
import pandas as pd
from datetime import datetime

# 输入文件路径（可修改）
input_file = r"C:\Users\Administrator\Downloads\product_20250903053750641.xls"

# 输出文件夹
output_folder = r"C:\Users\Administrator\Desktop\upload_sku"

# 获取当天日期 (格式：250903)
today_str = datetime.today().strftime("%y%m%d")

# 输出文件名
output_file = os.path.join(output_folder, f"占坑sku{today_str}.xlsx")

# 读取第一个子表，只取第一列（并去掉第一行标题）
df = pd.read_excel(input_file, sheet_name=0, usecols=[0], header=None)
df = df.iloc[1:, :]  # 删除第一行

# 创建新DataFrame
new_df = pd.DataFrame()
new_df["系统SKU"] = df.iloc[:, 0].values

# 生成第二列（加前缀）
prefix = f"ZLDBZHANK-{today_str}-01-"
new_df["平台SKU"] = prefix + new_df["系统SKU"].astype(str)

# 保存新文件
os.makedirs(output_folder, exist_ok=True)
new_df.to_excel(output_file, index=False)

print(f"处理完成 ✅ 新文件已保存：{output_file}")
