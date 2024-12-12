import pandas as pd
from collections import Counter
import re

import pandas as pd
import os
import subprocess


# 读取xlsx文件
df = pd.read_excel(r'C:\Users\123\Desktop\工作簿8.xlsx')

# 获取第一列数据
data = df.iloc[:, 0]

# 将所有词语拼接成一个长字符串
text = ' '.join(data)

# 使用正则表达式匹配连续的一个或两个或三个词语
one_word_matches = re.findall(r'\b\w+\b', text)
two_word_matches = re.findall(r'\b\w+\s\w+\b', text)
three_word_matches = re.findall(r'\b\w+\s\w+\s\w+\b', text)

# 使用Counter统计词频
one_word_count = Counter(one_word_matches)
two_word_count = Counter(two_word_matches)
three_word_count = Counter(three_word_matches)

# 根据词频进行排序
sorted_one_word_count = sorted(one_word_count.items(), key=lambda x: x[1], reverse=True)
sorted_two_word_count = sorted(two_word_count.items(), key=lambda x: x[1], reverse=True)
sorted_three_word_count = sorted(three_word_count.items(), key=lambda x: x[1], reverse=True)

# 提取最多出现的单个词、两个词和三个词
max_length = max(len(sorted_one_word_count), len(sorted_two_word_count), len(sorted_three_word_count))

# 创建一个空的DataFrame
result_df = pd.DataFrame(columns=['One Word', 'One Word Frequency', 'Two Words', 'Two Words Frequency', 'Three Words', 'Three Words Frequency'])

# 将词频数据填充到DataFrame中
for i in range(max_length):
    if i < len(sorted_one_word_count):
        result_df.loc[i, 'One Word'] = sorted_one_word_count[i][0]
        result_df.loc[i, 'One Word Frequency'] = sorted_one_word_count[i][1]
    if i < len(sorted_two_word_count):
        result_df.loc[i, 'Two Words'] = sorted_two_word_count[i][0]
        result_df.loc[i, 'Two Words Frequency'] = sorted_two_word_count[i][1]
    if i < len(sorted_three_word_count):
        result_df.loc[i, 'Three Words'] = sorted_three_word_count[i][0]
        result_df.loc[i, 'Three Words Frequency'] = sorted_three_word_count[i][1]

# 将结果写入新的xlsx文件



out_path=rf'C:\Users\123\Desktop\词频.xlsx'

result_df.to_excel(out_path, index=False)
# 调用函数并保存结果

if os.name == 'nt':  # Windows系统
    # 假设WPS的路径是 "C:\Program Files\WPS Office\et.exe"
    wps_path = r'C:\Users\123\AppData\Local\Kingsoft\WPS Office\ksolaunch.exe'
    # 使用WPS打开Excel文件
    subprocess.call([wps_path, out_path])
