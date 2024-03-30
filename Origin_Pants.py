import pandas as pd
import os
import numpy as np
from datetime import date

today = date.today()
date_string = today.strftime("%Y/%m/%d")
sku_date=today.strftime("%Y%m%d")

# 读取源表格文件
source_file = r'C:\Users\123\Desktop\product_01.xls'  # 替换为实际的源文件路径
df_source = pd.read_excel(source_file)

# 获取表格文件的名称
file_name = os.path.basename(source_file)
table_name = os.path.splitext(file_name)[0]

prefix = "product_"
new_string = table_name[len(prefix):]
print(new_string)

# 提取sku
sku_old = df_source.iloc[0:, 0]

sku = []
prefix = "meilu"+sku_date+new_string

for string in sku_old:
    new_string = prefix + string
    sku.append(new_string)

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
price = []

for string in price_str:
    start_index = string.find("*US灵境：") + len("*US灵境：")
    end_index = string.find(" USD  *US顺丰")
    new_string = string[start_index:end_index]
    rounded_float = round(float(new_string), 1) - 0.01
    formatted_string = "{:.2f}".format(rounded_float)
    price.append(formatted_string)

#提取图片
main_img = df_source.iloc[0:,11]
img_1 = df_source.iloc[0:,9]
img_2 = df_source.iloc[0:,15]
img_3 = df_source.iloc[0:,17]
img_4 = df_source.iloc[0:,19]
img_5 = df_source.iloc[0:,21]
img_6 = df_source.iloc[0:,23]
img_7 = df_source.iloc[0:,25]
final_img = df_source.iloc[0:,13]

# 创建新的数据框架
df_new = pd.DataFrame()

length = len(sku)  # 填充的长度
df_new['Column1'] = ['pants'] * length

# 插入目标数据到第二列
df_new['Column2'] = sku

#改牌子
brand = 'xxx'

df_new['brand'] = [brand] * length
df_new['update'] = ['Update'] * length
df_new['item_name'] = title

# 填充空白列
df_new[[f'blank{i}' for i in range(1, 3)]] = np.nan

df_new['product_description'] = product_description
string = title[0]
if "Casual" in string:
    item_type = 'casual-pants'
elif "Yoga" in string:
    item_type = 'yoga-pants'
else:
    item_type = 'pants'
df_new['item_type'] = [item_type] * length
string = title[0]
if "Drawstring" in string:
    closure_type = 'drawstring'
elif "Button Fly" in string:
    closure_type = 'button fly'
else:
    closure_type = 'elastic'
df_new['closure_type'] = [closure_type] * length

# 填充空白列
df_new[[f'blank{i}' for i in range(3, 14)]] = np.nan

df_new['price'] = price  # 后续要去除第一行
df_new['quantity'] = ['500'] * length
df_new['gender'] = ['Female'] * length
df_new['age'] = ['Adult'] * length
df_new['size_sys'] = ['US'] * length
df_new['Alpha'] = ['Alpha'] * length
df_new['size'] = size

# 填充空白列
df_new[[f'blank{i}' for i in range(18, 21)]] = np.nan

df_new['body_type'] = ['Regular'] * length
df_new['height_type'] = ['Regular'] * length
df_new['body_type'] = ['Regular'] * length
df_new['height_type'] = ['Regular'] * length
df_new['main_img'] = main_img
df_new['img_1'] = img_1
df_new['img_2'] = img_2
df_new['img_3'] = img_3
df_new['img_4'] = img_4
df_new['img_5'] = img_5
df_new['img_6'] = img_6
df_new['img_7'] = img_7
df_new['final_img'] = final_img
df_new['swatch_img'] = main_img
# 清除第1行数据
df_new.loc[0, ['price', 'quantity', 'age', 'Alpha', 'size_sys', 'body_type', 'height_type', 'main_img', 'img_1', 'img_2', 'img_3', 'img_4', 'img_5', 'img_6', 'img_7', 'final_img','swatch_img']] = np.nan


df_new['parent_child'] = ['Child'] * length
df_new.loc[0, 'parent_child'] = 'Parent'
df_new['parent_sku'] = [sku[0]] * length
df_new.loc[0, 'parent_sku'] = np.nan
df_new['relationship'] = ['Variation'] * length
df_new.loc[0, 'relationship'] = np.nan
df_new['variation_theme'] = ['color-size'] * length

# 填充空白列
df_new[[f'blank{i}' for i in range(31, 41)]] = np.nan

df_new['color'] = color
df_new['color_map'] = color
df_new.loc[0, ['color', 'color_map']] = np.nan

# 填充空白列
df_new[[f'blank{i}' for i in range(41, 42)]] = np.nan

df_new['department'] = ['Women'] * length
df_new.loc[0, 'department'] = np.nan
df_new['fit_type'] = ['Regular'] * length
df_new.loc[0, 'fit_type'] = np.nan
# 填充空白列
df_new[[f'blank{i}' for i in range(42, 47)]] = np.nan
df_new['rise_style'] = [rise] * length
df_new.loc[0, 'rise_style'] = np.nan
df_new[[f'blank{i}' for i in range(47, 58)]] = np.nan


df_new['is_autographed'] = ['No'] * length
df_new.loc[0, 'is_autographed'] = np.nan

# 填充空白列
df_new[[f'blank{i}' for i in range(58, 75)]] = np.nan
df_new['date'] = [date_string] * length
df_new.loc[0, 'date'] = np.nan
df_new[[f'blank{i}' for i in range(75, 78)]] = np.nan
string = title[0]
if "Flare" in string:
    leg_style = 'Flared'
elif "Cropped" in string:
    leg_style = 'Cropped'
else:
    leg_style = 'Wide'
df_new['leg_style'] = [leg_style] * length
df_new.loc[0, 'leg_style'] = np.nan
# 填充空白列
df_new[[f'blank{i}' for i in range(78, 91)]] = np.nan
# 创建一个新的数组，将修改后的字符串存储其中
size_map = [str(string).replace("2X", "XX").replace("3X", "XXX").replace("4X", "XXXX").replace("5X", "XXXXX") for string in size]

df_new['size_map'] = size_map
df_new.loc[0, 'size_map'] = np.nan

# 填充空白列
df_new[[f'blank{i}' for i in range(92, 112)]] = np.nan
if "jeans" in string:
    fabric_type = 'denim'
elif "linen" in string:
    fabric_type = 'linen'
elif "chiffon" in string:
    fabric_type = 'chiffon'
else:
    fabric_type = '71%Polyester,18%Cotton,11%Spandex'
df_new['fabric_type'] = [fabric_type] * length

# 填充空白列
df_new[[f'blank{i}' for i in range(112, 163)]] = np.nan

df_new['list_price'] = price
df_new.loc[0, 'list_price'] = np.nan
df_new.loc[0, [f'blank{i}' for i in range(163, 164)]] = np.nan
df_new['currency'] = ['USD'] * length
df_new.loc[0, 'currency'] = np.nan
df_new['condition_type'] = ['New'] * length
df_new.loc[0, 'condition_type'] = np.nan
# 填充空白列
df_new[[f'blank{i}' for i in range(164, 181)]] = np.nan


#改运费
shipping = 'xxx'


df_new['shipping'] = [shipping] * length
df_new.loc[0, 'shipping'] = np.nan

# 保存新的数据框架为 Excel 文件
output_filename = os.path.splitext(os.path.basename(source_file))[0] + '.xlsx'
output_file = os.path.join(output_filename)
df_new.to_excel(output_file, index=False)

print("新的目标表格已保存。")
