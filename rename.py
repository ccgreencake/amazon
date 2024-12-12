import os
import 复制 as copy

copy.copyfile(5)

folder_path = r"C:\Users\123\Desktop\240618"  # 替换为实际的文件夹路径
file_list = os.listdir(folder_path)
account = "li"
folder_name = os.path.basename(folder_path)
for i, file_name in enumerate(file_list):
    file_path = os.path.join(folder_path, file_name)
    if i < 9:
        new_file_name = account+folder_name + "0" + str(i + 1) + os.path.splitext(file_name)[1]
        print(new_file_name)
    else:
        new_file_name = account+folder_name + str(i + 1) + os.path.splitext(file_name)[1]
        print(new_file_name)
    new_file_path = os.path.join(folder_path, new_file_name)
    os.rename(file_path, new_file_path)