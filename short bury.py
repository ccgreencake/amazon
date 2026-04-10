import pandas as pd
import os
from datetime import date

today = date.today()
date_string = today.strftime("%Y/%m/%d")
sku_date=today.strftime("%Y%m%d")

# 读取源表格文件
source_file = r'C:\Users\123\Desktop\product\240424\product_08.xls'  # 替换为实际的源文件路径
df_source = pd.read_excel(source_file)


source_file2 = r'C:\Users\123\Desktop\new 词表.xlsx'  # 替换为实际的源文件路径
df_source2 = pd.read_excel(source_file2, sheet_name='yoga')

#提取埋词
word = df_source2.iloc[0:, 0]
# 创建新的数据框架
df_new = pd.DataFrame()
sku = df_source.iloc[0:, 0]
length = len(sku)  # 填充的长度
num = length*32
print("共需要",num,"个词")
print("词表长度为",len(word))
if num > len(word):
    print("词表不够，进行自动扩充")
while num > len(word):
    word = pd.concat([word, word], ignore_index=True)
    print("词表扩充中...")
    if num < len(word):
        word = word[:num]
        print("词表扩充完毕")
        break
    else:
        continue

for i in range(24):
    array_name = f"Array{i+1}"  # 创建数组名称
    start_index = i * length  # 计算起始索引
    end_index = (i + 1) * length  # 计算结束索引
    array_data = word[start_index:end_index].reset_index(drop=True)  # 提取32个元素并重置索引
    df_new[array_name] = array_data  # 将数据填充到新的数据框架的不同列中


# 保存新的数据框架为 Excel 文件
output_filename = os.path.splitext(os.path.basename(source_file))[0] + '.xlsx'
output_file = os.path.join(output_filename)
df_new.to_excel(output_file, index=False)

print("新的目标表格已保存。")