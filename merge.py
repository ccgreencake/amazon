import pandas as pd
import os

# 指定包含xlsx文件的文件夹路径
folder_path = 'C:\\Users\\123\\Desktop\\test'

# 收集所有xlsx文件的路径
xlsx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# 创建一个空的DataFrame，用于存储合并后的数据
merged_df = pd.DataFrame()

# 遍历所有xlsx文件并按行合并
for file in xlsx_files:
    # 读取xlsx文件
    df = pd.read_excel(file)

    # 获取文件名（不含扩展名）
    file_name = os.path.splitext(os.path.basename(file))[0]

    # 获取文件名的第五个字符到第11个字符
    time_str = file_name[5:11]

    # 添加时间列，列名为'Time'，内容为文件名的指定部分，并放在第一列
    df.insert(0, 'Time', time_str)

    # 将读取的数据追加到merged_df中
    merged_df = pd.concat([merged_df, df], ignore_index=True)

# 指定合并后的文件名
output_file = 'merged_file.xlsx'

# 将合并后的数据写入新的xlsx文件
merged_df.to_excel(output_file, index=False)

print(f'All files have been merged into {output_file}')