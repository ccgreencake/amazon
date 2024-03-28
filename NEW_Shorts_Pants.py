import pandas as pd
import os
import numpy as np
from datetime import date

today = date.today()
date_string = today.strftime("%Y/%m/%d")
sku_date=today.strftime("%Y%m%d")

# 读取源表格文件
source_file = r'C:\Users\123\Desktop\product\240328\product_07.xls'  # 替换为实际的源文件路径
df_source = pd.read_excel(source_file)


# 读取埋词文件
source_file2 = r'C:\Users\123\Desktop\埋词\shorts.xlsx'  # 替换为实际的源文件路径
df_source2 = pd.read_excel(source_file2)

#提取埋词
word = df_source2.iloc[0:, 0]
point5 = df_source2.iloc[0:, 1]


# 获取表格文件的名称
file_name = os.path.basename(source_file)
table_name = os.path.splitext(file_name)[0]

prefix = "product_"
new_string = table_name[len(prefix):]

# 提取sku
sku_old = df_source.iloc[0:, 0]

sku = []
prefix = "meilu"+sku_date+new_string

for string in sku_old:
    new_string = prefix + str(string)
    sku.append(new_string)

# 提取title
title = df_source.iloc[0:, 2]
string = title[0]
if "High" in string:
    rise = 'high'
elif "Low" in string:
    rise = 'low'
else:
    rise = 'mid'

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

df_new['brand'] = ['sttsgwyt'] * length
df_new['update'] = ['Update'] * length
df_new['item_name'] = title



size_chart = df_source.iloc[1,26]
p5 = df_source.iloc[0:,27]
part1 = p5[0].split(":")
part2 = p5[1].split(":")
part3 = p5[2].split(":")
df_new['p1'] = [p5[0]]*length
df_new['p2'] = [p5[1]]*length
df_new['p3'] = [p5[2]]*length
df_new['p4'] = [p5[3]]*length
output_string = size_chart.replace("Size: ", "").replace(" Waist: ", " &emsp; ").replace(" Hip: ", " &emsp; ").replace(" Length: ", " &emsp; ").replace("\n", " <br>\n")
# 打开文本文件进行读取
with open('html.txt', 'r',encoding='utf-8') as file:
    # 读取文件内容
    file_content = file.read()
# 进行替换操作
file_content = file_content.replace("<strong>Length</strong><br>","<strong>Length</strong><br>\n"+output_string)
first_position = file_content.find("<p><strong></strong></p>")
# 查找第二个 <p><strong> 的位置
second_position = file_content.find("<p><strong></strong></p>", first_position + 1)
# 查找第3个 <p><strong> 的位置
third_position = file_content.find("<p><strong></strong></p>", second_position + 1)
# 查找第一个 <p></p> 的位置
first_p = file_content.find("<p></p>")
# 查找第二个 <p></p>  的位置
second_p = file_content.find("<p></p>", first_p + 1)
# 查找第二个 <p></p>  的位置
third_p = file_content.find("<p></p>", second_p + 1)


product_description = file_content[:third_p] + "<p>"+part3[1]+"</p>" + file_content[third_p+7:]
product_description = product_description[:third_position]+"<p><strong>"+part3[0]+"</strong></p>"+product_description[third_position+24:]
product_description = product_description[:second_p] + "<p>"+part2[1]+"</p>" + product_description[second_p+7:]
product_description = product_description[:second_position] + "<p><strong>"+part2[0]+"</strong></p>" + product_description[second_position+24:]
product_description = product_description[:first_p] + "<p>"+part1[1]+"</p>" + product_description[first_p+7:]
product_description = product_description[:first_position] + "<p><strong>"+part1[0]+"</strong></p>" + product_description[first_position+24:]

df_new['product_description'] = product_description
print('大描述长度：'+str(len(product_description)))
string = title[0]
if "Casual" in string:
    item_type = 'casual-pants'
elif "Yoga" in string:
    item_type = 'yoga-pants'
else:
    item_type = 'pants'
df_new['item_type'] = [item_type] * length
n = 1
word_name = 'word'+str(n)
df_new[word_name] = [word[n]] * length
n=n+1
for i in range(5):
    word_name = 'word'+str(n)
    df_new[word_name] = [word[n]] * length
    n=n+1
df_new['price'] = price  # 后续要去除第一行
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

for i in range(5):
    name = 'point'+str(i)
    df_new[name] = [point5[i]] * length
key = title[1] + ' ' + word[1]+' ' +  word[2]+' ' +  word[3]

df_new['key'] = [key.lower()] * length
print('key:'+str(len(key)))

df_new['color'] = new_color
df_new['color_map'] = new_color
df_new.loc[0, ['color', 'color_map']] = np.nan



df_new['department'] = ['Women'] * length
df_new.loc[0, 'department'] = np.nan
df_new['fit_type'] = ['Regular'] * length
df_new.loc[0, 'fit_type'] = np.nan



df_new['rise_style'] = [rise] * length
df_new.loc[0, 'rise_style'] = np.nan



df_new['is_autographed'] = ['No'] * length
df_new.loc[0, 'is_autographed'] = np.nan

for i in range(5):
    word_name = 'word'+str(n)
    df_new[word_name] = [word[n]] * length
    n=n+1



word_name = 'word'+str(n)
df_new[word_name] = [word[n]] * length
n=n+1



word_name = 'word'+str(n)
df_new[word_name] = [word[n]] * length
n=n+1

string = title[0]
if "Flare" in string:
    leg_style = 'Flared'
elif "Cropped" in string:
    leg_style = 'Cropped'
else:
    leg_style = 'Wide'
df_new['leg_style'] = [leg_style] * length
df_new.loc[0, 'leg_style'] = np.nan


for i in range(3):
    word_name = 'word'+str(n)
    df_new[word_name] = [word[n]] * length
    n=n+1



# 创建一个新的数组，将修改后的字符串存储其中
size_map = [str(string).replace("2X", "XX").replace("3X", "XXX").replace("4X", "XXXX").replace("5X", "XXXXX") for string in size]

df_new['size_map'] = size_map
df_new.loc[0, 'size_map'] = np.nan
df_new['size_name'] = size_map
df_new.loc[0, 'size_name'] = np.nan
word_name = 'word'+str(n)
df_new[word_name] = [word[n]] * length
n=n+1
word_name = 'word'+str(n)
df_new[word_name] = [word[n]] * length
n=n+1
word_name = 'word'+str(n)
df_new[word_name] = [word[n]] * length
n=n+1
word_name = 'word'+str(n)
df_new[word_name] = [word[n]] * length
n=n+1
df_new['list_price'] = price
df_new.loc[0, 'list_price'] = np.nan
df_new['currency'] = ['USD'] * length
df_new.loc[0, 'currency'] = np.nan
df_new['condition_type'] = ['New'] * length
df_new.loc[0, 'condition_type'] = np.nan


df_new['shipping'] = ['Jeweli'] * length
df_new.loc[0, 'shipping'] = np.nan

# 保存新的数据框架为 Excel 文件
output_filename = os.path.splitext(os.path.basename(source_file))[0] + '.xlsx'
output_file = os.path.join(output_filename)
df_new.to_excel(output_file, index=False)

print("新的目标表格已保存。")
