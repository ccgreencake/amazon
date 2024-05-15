import pandas as pd
import os

from datetime import date

def build_sku(account, folder_path):
    today = date.today()
    sku_date = today.strftime("%Y%m%d")



    # 将 folder_path 末尾的路径分隔符去除
    if folder_path.endswith(os.sep):
        folder_path = folder_path[:-1]

    # 构建新的文件夹路径，将 account 变量的值添加到末尾
    folder_path = folder_path + os.sep + account # 替换为实际的文件夹路径
    file_list = os.listdir(folder_path)
    # 获取上一级文件夹的名字
    folder_name = os.path.basename(os.path.dirname(folder_path))


    all_sku = []
    my_sku = []
    for i, file_name in enumerate(file_list):
        if '~' in file_name:
            continue
        file_path = os.path.join(folder_path, file_name)
        # 读取源表格文件
        source_file = file_path  # 替换为实际的源文件路径
        # 获取表格文件的名称
        file_name = os.path.basename(source_file)
        table_name = os.path.splitext(file_name)[0]

        prefix = "product_"
        new_string = table_name[len(prefix):]
        df_source = pd.read_excel(source_file)
        # 获取表格文件的名称
        file_name = os.path.basename(source_file)
        table_name = os.path.splitext(file_name)[0]



        # 提取sku
        sku_old = df_source.iloc[:, 0].values
        sku = ["T"+account + "20" + folder_name + new_string + str(string) for string in sku_old]
        all_sku.extend(sku_old)
        my_sku.extend(sku)

    # 创建新的数据框架
    df_new = pd.DataFrame()
    df_new['SKU'] = all_sku
    df_new['平台SKU'] = my_sku

    # 构建新的 name 变量
    name = "sku20" + folder_name

    output_folder = "C:\\Users\\123\\Desktop\\upload_sku"  # 设置保存的文件夹路径

    # 构建完整的文件路径
    output_filename = os.path.join(output_folder, os.path.splitext(os.path.basename(name))[0] + account + '.xlsx')

    # 保存数据框架为 Excel 文件
    df_new.to_excel(output_filename, index=False)

    print("新的目标表格已保存。")
    return output_filename


def merge(name1,name2):

    # 读取第一个 Excel 文件
    df1 = pd.read_excel(name1)

    # 读取第二个 Excel 文件
    df2 = pd.read_excel(name2)

    # 合并两个数据框
    merged_df = pd.concat([df1, df2], ignore_index=True)
    merge_name = name2.replace('li', '_merge')

    # 写入合并后的数据框到新的 Excel 文件
    merged_df.to_excel(merge_name, index=False)
    print(f"合并后的文件已保存为 '{merge_name}'。")


    # 检查文件是否存在
    if os.path.exists(name1):
        # 删除文件
        os.remove(name1)
        print(f"文件 '{name1}' 已删除。")
    else:
        print(f"文件 '{name1}' 不存在。")
    # 检查文件是否存在
    if os.path.exists(name2):
        # 删除文件
        os.remove(name2)
        print(f"文件 '{name2}' 已删除。")
    else:
        print(f"文件 '{name2}' 不存在。")


folder_path = r"C:\Users\123\Desktop\product\240514"
# name1 = build_sku("xh",folder_path)
name2 = build_sku("li",folder_path)
# merge(name1,name2)