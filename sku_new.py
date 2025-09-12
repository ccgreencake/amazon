import os
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import date

# 文件夹路径
folder_path = "C:\\Users\\Administrator\\Desktop\\250911"

# 获取今天的日期
today = date.today()
date_str = today.strftime("%Y%m%d")

folder_name = os.path.basename(folder_path)
print(folder_name)
# 新建 .xlsx 文件路径
xlsx_file_path = f"C:\\Users\\Administrator\\Desktop\\upload_sku\\sku{folder_name}.xlsx"

# 创建一个新的 Workbook 对象
wb_dest = Workbook()

# 添加一个工作表到新的 Workbook
ws_dest = wb_dest.active
ws_dest.title = 'Merged Data'

# 添加标题行
ws_dest["A1"] = "系统SKU"
ws_dest["B1"] = "平台SKU"

# 修改前的数据保存到数组
original_data = []

# 修改后的数据保存到数组
processed_data = []

# 遍历文件夹中的所有文件
for filename in os.listdir(folder_path):
    if filename.startswith('~$'):
        continue
    if filename.endswith('.xlsx'):
        # 拼接文件的完整路径
        file_path = os.path.join(folder_path, filename)

        # 加载 .xlsm 文件，使用只读模式
        wb_src = load_workbook(filename=file_path, read_only=True)

        # 选择工作表
        if "sa" in filename:
            ws_src = wb_src['Vorlage']
        else:
            ws_src = wb_src['Template']

        # 遍历 .xlsm 文件中的第二列（从第四行开始）
        for row in ws_src.iter_rows(min_row=4, max_col=2, values_only=True):
            # 如果遇到空值（None），跳过该值的读取
            if row and row[1] is not None:
                sku = row[1]
                original_data.append(sku)  # 保存修改前的数据
                if sku.startswith('V') or sku.startswith('T') or sku.startswith('C') or sku.startswith('Z') or sku.startswith('P'):
                    sku = sku[20:]
                elif sku.startswith('li') or sku.startswith('yo') or sku.startswith('lc'):
                    sku = sku[10:]
                processed_data.append(sku)  # 保存修改后的数据
                ws_dest.append([sku, sku])

        # 关闭当前文件的工作簿
        wb_src.close()

# 将修改前的数据保存到第一列
for i, sku in enumerate(original_data):
    ws_dest.cell(row=i+2, column=2, value=sku)

# 将修改后的数据保存到第二列
for i, sku in enumerate(processed_data):
    ws_dest.cell(row=i+2, column=1, value=sku)

# 保存新的 .xlsx 文件
wb_dest.save(filename=xlsx_file_path)

# 关闭工作簿
wb_dest.close()