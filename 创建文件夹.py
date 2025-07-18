import os
from datetime import datetime


def get_multiline_input():
    print("请输入需要创建的文件夹名称（每行一个，输入空行结束）：")
    lines = []
    while True:
        line = input().strip()
        if not line:  # 输入空行时结束
            break
        lines.append(line)
    return lines


def create_folders():
    # 获取文件夹名称列表
    folder_names = get_multiline_input()

    # 固定目标路径
    target_directory = r"F:\产品归类"

    if not folder_names:
        print("未输入有效的文件夹名称！")
        return

    # 确保目标目录存在
    if not os.path.exists(target_directory):
        print(f"目标目录 {target_directory} 不存在，正在创建...")
        os.makedirs(target_directory)

    # 遍历创建文件夹
    for folder_name in folder_names:
        folder_path = os.path.join(target_directory, folder_name)

        if os.path.exists(folder_path):
            print(f"\n文件夹 '{folder_name}' 已存在，正在创建时间戳文件...")
            try:
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                txt_file_name = f"{current_time}.txt"
                txt_file_path = os.path.join(folder_path, txt_file_name)
                with open(txt_file_path, 'w') as f:
                    f.write(f"Created at {datetime.now()}")
                print(f"已创建时间戳文件：{txt_file_path}")
            except Exception as e:
                print(f"创建时间戳文件失败：{e}")
        else:
            try:
                os.makedirs(folder_path)
                print(f"\n已成功创建文件夹：{folder_path}")
                # 在新建文件夹中创建初始文件
                initial_file = os.path.join(folder_path, "folder_created.txt")
                with open(initial_file, 'w') as f:
                    f.write(f"Created at {datetime.now()}")
            except Exception as e:
                print(f"创建文件夹失败：{folder_path}\n错误信息：{e}")


if __name__ == "__main__":
    print("=" * 40)
    print("智能文件夹创建工具")
    print("=" * 40)
    create_folders()
    print("\n操作已完成，请按回车键退出...")
    input()
