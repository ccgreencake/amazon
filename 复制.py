import os
import shutil
from datetime import datetime


def copyfile1(n,account,path):
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


def copyfile2(file_path,n):
    # 检查文件是否存在
    if not os.path.isfile(file_path):
        print(f"文件不存在: {file_path}")
        return

    # 获取文件所在的目录和文件名（不含扩展名）
    directory, filename = os.path.split(file_path)
    name, ext = os.path.splitext(filename)

    # 确保扩展名为 .xls 或 .xlsx
    if ext.lower() not in ['.xls', '.xlsx']:
        print("文件必须是 .xls 或 .xlsx 格式。")
        return

    # 检查 n 是否为正整数
    if not isinstance(n, int) or n <= 0:
        print("n 必须是一个正整数。")
        return

    # 复制文件 n 次
    for i in range(1, n + 1):
        # 格式化数字，例如 1 -> 01, 10 -> 10
        num_str = f"{i:02d}"
        new_filename = f"product_{num_str}{ext}"
        new_file_path = os.path.join(directory, new_filename)

        # 检查文件是否已存在，避免覆盖
        if os.path.exists(new_file_path):
            print(f"文件已存在，跳过: {new_file_path}")
            continue

        try:
            shutil.copy2(file_path, new_file_path)
            print(f"已创建: {new_file_path}")
        except Exception as e:
            print(f"复制文件时出错: {e}")


def copyfile3(n, account, path):
    # 根据账户名映射到简短形式
    if account == "YOULE":
        account = 'yo'
    elif account == "LIANG":
        account = 'li'
    elif account == "SANSK":
        account = 'sa'
    elif account == "SIYAT":
        account = "si"

    # 获取当前日期并格式化为 YYMMDD 的形式
    folder_name = datetime.now().strftime('%y%m%d')
    # 设置桌面路径
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    # 创建文件夹的完整路径
    folder_path = os.path.join(desktop_path, folder_name)

    # 检查文件夹是否已存在，如果不存在则创建
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # 源文件路径
    source_file_path = os.path.join(desktop_path, path)

    # 构造新文件名并复制文件
    new_file_name = f'总模板_{n}.xlsx'
    new_file_path = os.path.join(folder_path, new_file_name)
    shutil.copy2(source_file_path, new_file_path)

    # 输出复制成功的信息
    print(f'文件 {new_file_name} 已复制到 {folder_path}')

    # 生成重命名的新文件名，只对刚刚复制的文件进行重命名
    # 提取文件扩展名
    extension = os.path.splitext(new_file_name)[1]

    # 定义新的重命名文件名，按规则加前导零或不加
    if n < 10:
        renamed_file_name = f'{account}{folder_name}0{n}{extension}'
    else:
        renamed_file_name = f'{account}{folder_name}{n}{extension}'

    renamed_file_path = os.path.join(folder_path, renamed_file_name)

    # 将文件重命名
    os.rename(new_file_path, renamed_file_path)
    print(f'文件已重命名为 {renamed_file_name}')

#
# n = 12
# account = "yo"
# # account = "li"
# # account = "zz"
# if account == "yo":
#     path = 'yo总模板.xlsx'
# elif account == "li":
#     path = 'li总模板.xlsx'
# elif account == "zz":
#     path = 'zz模板.xlsx'
#
# # copyfile1(n,account,path)
# path = r"C:\Users\Administrator\Desktop\product\250821\product_04.xls"
# copyfile2(path,n)