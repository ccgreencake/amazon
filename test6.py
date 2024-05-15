import pandas as pd

# 读取第一个 Excel 文件
df1 = pd.read_excel("C:\\Users\\123\\Desktop\\upload_sku\\sku20240501li.xlsx")

# 读取第二个 Excel 文件
df2 = pd.read_excel("C:\\Users\\123\\Desktop\\upload_sku\\sku20240501xh.xlsx")

# 合并两个数据框
merged_df = pd.concat([df1, df2], ignore_index=True)

# 写入合并后的数据框到新的 Excel 文件
merged_df.to_excel('C:\\Users\\123\\Desktop\\upload_sku\\', index=False)