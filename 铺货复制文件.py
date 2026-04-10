import os
import shutil
import re


def batch_copy_files(folder_path, file_name, num_copies):
    # 1. 拼接完整源文件路径
    source_path = os.path.join(folder_path, file_name)

    # 检查源文件是否存在
    if not os.path.exists(source_path):
        print(f"错误: 找不到文件 {source_path}")
        return

    # 2. 分离文件名和后缀 (例如: 'product_04' 和 '.xls')
    name_part, extension = os.path.splitext(file_name)

    # 3. 使用正则表达式提取文件名末尾的数字及其位数
    # 查找末尾的数字，例如 "product_04" 会匹配到 "04"
    match = re.search(r'(\d+)$', name_part)

    if match:
        num_str = match.group(1)  # 提取数字字符串 "04"
        start_num = int(num_str)  # 转为整数 4
        width = len(num_str)  # 获取数字位数（用于补0，保持格式）
        prefix = name_part[:match.start()]  # 获取数字前的部分 "product_"
    else:
        # 如果文件名末尾没有数字，默认从1开始，并在中间加下划线
        prefix = name_part + "_"
        start_num = 0
        width = 2

    print(f"开始复制文件，源文件: {file_name}")

    # 4. 循环复制
    for i in range(1, num_copies + 1):
        next_num = start_num + i
        # 格式化新文件名，zfill(width) 用于补齐前导0
        new_file_name = f"{prefix}{str(next_num).zfill(width)}{extension}"
        dest_path = os.path.join(folder_path, new_file_name)

        # 执行复制
        shutil.copy2(source_path, dest_path)
        print(f"已生成: {new_file_name}")

    print("--- 复制任务完成 ---")


# ================= 配置区域 =================
# 目标路径 (建议在路径前加 r 防止转义)
target_path = r'C:\Users\Administrator\Desktop\product\260108'

# 目标文件名
target_file = 'product_05.xls'

# 复制个数
copy_count = 10
# ===========================================

# 执行脚本
if __name__ == "__main__":
    batch_copy_files(target_path, target_file, copy_count)