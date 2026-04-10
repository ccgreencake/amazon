import os
import pandas as pd
from datetime import datetime


def merge_xlsx_with_monthly_increment():
    # ================= 配置区域 =================
    # 1. 设置文件夹路径 (请修改为你实际的路径)
    folder_path = r"C:\Users\Administrator\Desktop\order"

    # 设置输出文件名称
    output_file_name = "合并后的Excel数据.xlsx"

    # 2. 定义起始时间 (2024年6月)
    start_year = 2023
    start_month = 11
    # ===========================================

    output_path = os.path.join(folder_path, output_file_name)

    # 3. 获取并筛选文件
    all_files = os.listdir(folder_path)

    files_to_process = []

    for filename in all_files:
        # 跳过输出文件自己，防止死循环
        if filename == output_file_name:
            continue

        # 筛选 .xlsx 文件，且不以 ~$ 开头（忽略Excel打开时的临时缓存文件）
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            files_to_process.append(filename)

    # 4. 排序：按照文本名称递增排序
    # Python 默认的 sort 就是文本（字典）顺序
    files_to_process.sort()

    if not files_to_process:
        print("未找到 .xlsx 文件。")
        return

    print(f"共找到 {len(files_to_process)} 个文件，按名称排序如下：")
    print(files_to_process)
    print("-" * 30)

    all_dataframes = []

    # 5. 循环处理
    for index, filename in enumerate(files_to_process):
        file_path = os.path.join(folder_path, filename)

        try:
            # --- 日期计算逻辑 (核心) ---
            # index 是 0, 1, 2...
            # 我们需要计算经过 index 个月后的年份和月份

            # 先计算从起始点开始的总月数
            total_months = (start_year * 12 + start_month - 1) + index

            # 反算出当前的年和月
            current_year = total_months // 12
            current_month = (total_months % 12) + 1

            # 格式化为 "2024.06" 这种字符串
            # :02d 表示月份如果是个位数，前面补0
            date_str = f"{current_year}{current_month:02d}"

            # --- 读取 Excel ---
            # 默认 header=0，即第一行是标题
            df = pd.read_excel(file_path)

            # --- 插入时间列 ---
            # 在第0列（最左边）插入
            df.insert(0, '时间', date_str)

            all_dataframes.append(df)

            print(f"处理: {filename} -> 对应时间: {date_str}")

        except Exception as e:
            print(f"读取文件 {filename} 失败: {e}")

    # 6. 合并与保存
    if all_dataframes:
        print("-" * 30)
        print("正在合并，请稍候...")
        final_df = pd.concat(all_dataframes, ignore_index=True)

        # 保存为 Excel
        final_df.to_excel(output_path, index=False)

        print(f"成功！已合并 {len(all_dataframes)} 个文件。")
        print(f"结果保存在: {output_path}")
    else:
        print("没有数据被合并。")


if __name__ == "__main__":
    # 确保你安装了 pandas 和 openpyxl
    # pip install pandas openpyxl
    merge_xlsx_with_monthly_increment()