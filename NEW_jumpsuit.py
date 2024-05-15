import pandas as pd
import os
import numpy as np
from datetime import date
import utils

today = date.today()
date_string = today.strftime("%Y/%m/%d")
sku_date=today.strftime("%Y%m%d")

# 读取源表格文件
source_file = r'C:\Users\123\Desktop\product\240424\product_09.xls'  # 替换为实际的源文件路径
df_source = pd.read_excel(source_file)


source_file2 = r'C:\Users\123\Desktop\new 词表.xlsx'  # 替换为实际的源文件路径
df_source2 = pd.read_excel(source_file2, sheet_name='jumpsuit')

#提取埋词
word = df_source2.iloc[0:, 0]
point5 = df_source2.iloc[0:, 1]

price='19.99'
list_price='19.99'

# 获取表格文件的名称
file_name = os.path.basename(source_file)
table_name = os.path.splitext(file_name)[0]

prefix = "product_"
new_string = table_name[len(prefix):]

# 提取sku
sku_old = df_source.iloc[0:, 0]

sku = []
prefix = "li"+sku_date+new_string

for string in sku_old:
    new_string = prefix + str(string)
    sku.append(new_string)

# 提取title
title = df_source.iloc[0:, 2]


# 提取color
colors = df_source.iloc[0:, 4]
n=0
new_color = []

for i in range(len(colors)):
    if i != 0:
        if colors[i] != colors[(i-1)]:
            n = n+1
    prefix = "A" + str(n).zfill(3) + "-"
    color = prefix + str(colors[i])
    new_color.append(color)
# 提取size
size = df_source.iloc[0:, 5]
# # 提取价格
# price_str = df_source.iloc[0:, 7]
# price = []
# list_price=[]
#
# for string in price_str:
#     start_index = string.find("*US灵境：") + len("*US灵境：")
#     end_index = string.find(" USD  *US顺丰")
#     new_string = string[start_index:end_index]
#     rounded_float = round(float(new_string), 1) - 2.01
#     rounded_float1 = rounded_float+6
#     formatted_string = "{:.2f}".format(rounded_float)
#     formatted_string1 = "{:.2f}".format(rounded_float1)
#     price.append(formatted_string)
#     list_price.append(formatted_string1)

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
string = title[0]
length = len(sku)  # 填充的长度
feed_product = utils.feed_product(string)

df_new['Column1'] = [feed_product] * length

# 插入目标数据到第二列
df_new['Column2'] = sku

df_new['brand'] = ['Bakgeerle'] * length
df_new['update'] = ['Update'] * length
colors[0] = np.nan
new_title = [str(f_title) + ' ' + str(color) if i > 0 else f_title for i, (f_title, color) in enumerate(zip(title, colors))]
df_new['item_name'] = new_title

# 暂时放空
df_new['waist_style'] = np.nan
df_new['material_type1'] = np.nan
df_new['material_type2'] = np.nan
df_new['special_features1'] = np.nan



df_new['style_name'] = ['casual']*length
df_new['lifecycle_supply_type'] = ['Year Round Replenishable']*length


#暂时放空
df_new['neck_style'] = np.nan
df_new['sleeve_type'] = np.nan


df_new['cpsia_cautionary_statement'] = ['NoWarningApplicable']*length
df_new['country_of_origin'] = ['China']*length
df_new['supplier_declared_material_regulation1'] = ['Not Applicable']*length



df_new['length'] = ['regular']*length
df_new['key'] = np.nan



size_chart = df_source.iloc[1,26]
p5 = df_source.iloc[0:,27]
part1 = p5[0].split(":")
part2 = p5[1].split(":")
part3 = p5[2].split(":")
for i in range(4):
    # 找到【的索引位置
    index = p5[i].find('【')

    # 如果找到了【，则去除该字符及其之前的部分
    if index != -1:
        p5[i] = p5[i][index:]

df_new['p1'] = [p5[0]]*length
df_new['p2'] = [p5[1]]*length
if feed_product == 'shorts':
    if 'Linen' in string:
        df_new['p3'] = ['55%linen, 45%cotton'] * length
    else:
        df_new['p3'] = ['71%Polyester,18%Cotton,11%Spandex']*length
else:
    df_new['p3'] = [p5[2]] * length
df_new['p4'] = [p5[3]]*length

product_description = utils.description("html.txt", size_chart, part1, part2, part3)

n = 1
for i in range(17):
    word_name = 'word'+str(n)
    df_new[word_name] = [word[n]] * length
    n=n+1
df_new['product_description'] = product_description
print('大描述长度：'+str(len(product_description)))

item_type = utils.item_type(string,feed_product)
df_new['item_type'] = [item_type] * length

df_new['price'] = [price] * length  # 后续要去除第一行
df_new['quantity'] = ['500'] * length
df_new['gender'] = ['Female'] * length
df_new['age'] = ['Adult'] * length
df_new['size_sys'] = ['US'] * length
df_new['Alpha'] = ['Alpha'] * length
df_new['size'] = size


df_new['body_type'] = ['Regular'] * length
df_new['height_type'] = ['Regular'] * length
df_new['main_img'] = main_img
df_new['img_1'] = img_1
df_new['img_2'] = img_2
df_new['img_3'] = img_3
df_new['img_4'] = img_4
df_new['img_5'] = img_5
df_new['img_6'] = final_img
df_new['img_7'] = img_7
df_new['final_img'] = img_6
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
for i in range(5):
    name = 'point'+str(i)
    df_new[name] = [point5[i]] * length


df_new['color'] = new_color
df_new['color_map'] = new_color
df_new.loc[0, ['color', 'color_map']] = np.nan



df_new['department'] = ['Women'] * length
df_new.loc[0, 'department'] = np.nan
df_new['fit_type'] = ['Regular'] * length
df_new.loc[0, 'fit_type'] = np.nan







df_new['is_autographed'] = ['No'] * length
df_new.loc[0, 'is_autographed'] = np.nan





# 创建一个新的数组，将修改后的字符串存储其中
size_map = [str(string).replace("2X", "XX").replace("3X", "XXX").replace("4X", "XXXX").replace("5X", "XXXXX") for string in size]

df_new['size_map'] = size_map
df_new.loc[0, 'size_map'] = np.nan
df_new['size_name'] = size_map
df_new.loc[0, 'size_name'] = np.nan


df_new['list_price'] = [list_price] * length
df_new.loc[0, 'list_price'] = np.nan
df_new['currency'] = ['USD'] * length
df_new.loc[0, 'currency'] = np.nan
df_new['condition_type'] = ['New'] * length
df_new.loc[0, 'condition_type'] = np.nan


df_new['shipping'] = ['0'] * length
df_new.loc[0, 'shipping'] = np.nan

# 保存新的数据框架为 Excel 文件
output_filename = os.path.splitext(os.path.basename(source_file))[0] + '.xlsx'
output_file = os.path.join(output_filename)
df_new.to_excel(output_file, index=False)

print("新的目标表格已保存。")
