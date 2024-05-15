import os

file_path = "C:\\Users\\123\\Desktop\\upload_sku\\"+"sku20240501xh.xlsx"

# 检查文件是否存在
if os.path.exists(file_path):
    # 删除文件
    os.remove(file_path)
    print(f"文件 '{file_path}' 已删除。")
else:
    print(f"文件 '{file_path}' 不存在。")