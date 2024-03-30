import pandas as pd
import os
import numpy as np
from datetime import date

today = date.today()
sku_date=today.strftime("%Y%m%d")
folder_path = r"C:\Users\123\Desktop\product\240329"  # 替换为实际的文件夹路径
file_list = os.listdir(folder_path)
all_sku = []
my_sku = []
for i, file_name in enumerate(file_list):
    file_path = os.path.join(folder_path, file_name)
    # 读取源表格文件
    source_file = file_path  # 替换为实际的源文件路径
    df_source = pd.read_excel(source_file)
    # 获取表格文件的名称
    file_name = os.path.basename(source_file)
    table_name = os.path.splitext(file_name)[0]

    prefix = "meilu"


    # 提取sku
    sku_old = df_source.iloc[:, 0].values
    sku = [prefix + sku_date + str(string) for string in sku_old]
    all_sku.extend(sku_old)
    my_sku.extend(sku)


# 创建新的数据框架
df_new = pd.DataFrame()
df_new['SKU'] = all_sku
df_new['平台SKU'] = my_sku
name = 'sku'+sku_date
# 保存新的数据框架为 Excel 文件
output_filename = os.path.splitext(os.path.basename(name))[0] + '.xlsx'
output_file = os.path.join(output_filename)
df_new.to_excel(output_file, index=False)

print("新的目标表格已保存。")
