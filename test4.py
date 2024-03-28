import pandas as pd
import os
import numpy as np
from datetime import date
import re


# 读取源表格文件
source_file = r'C:\Users\123\Desktop\product\240328\product_03.xls'  # 替换为实际的源文件路径
df_source = pd.read_excel(source_file)



size_chart = df_source.iloc[1,26]
p5 = df_source.iloc[0:,27]
part1 = p5[0].split(":")
part2 = p5[1].split(":")
part3 = p5[2].split(":")

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

# 将替换后的内容写入新的文本文件
with open('output.txt', 'w',encoding='utf-8') as file:
    file.write(product_description)
# 提取size
size = df_source.iloc[0:, 5]

