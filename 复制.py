import os
import shutil
from datetime import datetime


def copyfile(n,account,path):
    # 获取当前日期并格式化为 YYYYMMDD 的形式
    folder_name = datetime.now().strftime('%y%m%d')
    # 设置桌面路径（这取决于你的操作系统和设置）
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    # 创建文件夹的完整路径
    folder_path = os.path.join(desktop_path, folder_name)

    # 检查文件夹是否已存在，如果不存在则创建
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # 源文件路径（你的 Excel 文件）
    source_file_path = os.path.join(desktop_path, path)


    # 复制文件到新文件夹
    for i in range(n):
        # 构造新文件名，这里简单地使用序号 i 作为文件名
        new_file_name = f'总模板_{i}.xlsx'
        # 完整的新文件路径
        new_file_path = os.path.join(folder_path, new_file_name)
        # 复制文件
        shutil.copy2(source_file_path, new_file_path)
        print(f'文件 {new_file_name} 已复制到 {folder_path}')

    print(f'所有文件已复制到 {folder_path}')
    file_list = os.listdir(folder_path)

    folder_name = os.path.basename(folder_path)
    for i, file_name in enumerate(file_list):
        file_path = os.path.join(folder_path, file_name)
        if i < 9:
            new_file_name = account + folder_name + "0" + str(i + 1) + os.path.splitext(file_name)[1]
            print(new_file_name)
        else:
            new_file_name = account + folder_name + str(i + 1) + os.path.splitext(file_name)[1]
            print(new_file_name)
        new_file_path = os.path.join(folder_path, new_file_name)
        os.rename(file_path, new_file_path)
n = 3
account = "yo"
# account = "li"
# account = "lc"
if account == "yo":
    path = 'yo总模板.xlsx'
elif account == "li":
    path = 'li总模板.xlsx'
elif account == "lc":
    path = 'lcy模板.xlsx'

copyfile(n,account,path)