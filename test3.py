import pandas as pd
import os
import numpy as np
from datetime import date

today = date.today()
date_string = today.strftime("%Y/%m/%d")

# 读取源表格文件
source_file = r'C:\Users\123\Desktop\product_24032301.xls'  # 替换为实际的源文件路径
df_source = pd.read_excel(source_file)

# 提取sku
sku = df_source.iloc[0:, 0]

# 提取title
title = df_source.iloc[0:, 2]
print(title[0])
string = title[0]
if "High" in string:
    rise = 'high'
elif "Low" in string:
    rise = 'low'
else:
    rise = 'mid'
# 提取description
product_description = df_source.iloc[0:, 3]
# 提取color
color = df_source.iloc[0:, 4]
# 提取size
size = df_source.iloc[0:, 5]
# 提取价格
price_str = df_source.iloc[0:, 7]
new_array = []

for string in price_str:
    start_index = string.find("*US灵境：") + len("*US灵境：")
    end_index = string.find(" USD  *US顺丰")
    new_string = string[start_index:end_index]
    rounded_float = round(float(new_string), 1) - 0.01
    formatted_string = "{:.2f}".format(rounded_float)
    new_array.append(formatted_string)

print(new_array)