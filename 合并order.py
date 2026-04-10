import os
import shutil
import subprocess
from datetime import datetime
import xlrd  # 用于读取 xls
import xlwt  # 用于写入 xls
from xlutils.copy import copy  # 用于将读取的对象转为可写的对象


def merge_xls(folder_path, header_rows=1, target_sheet="账号业务统计", use_grouping=True, exclude_merge=True):
    """
    :param exclude_merge: 如果为 True, 排除文件名中包含 'merge' 的文件
    """
    # 1. 获取文件夹内所有有效的 xls 文件
    raw_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.xls') and not f.startswith('~$')]

    # 2. 排序逻辑：按文件名升序排列
    raw_files.sort()

    all_files = []
    for f in raw_files:
        if exclude_merge and 'merge' in f.lower():
            continue
        all_files.append(f)

    if not all_files:
        print(f"文件夹内没有找到可处理的 xls 文件。")
        return

    folder_name = os.path.basename(os.path.abspath(folder_path))

    # 3. 分组逻辑
    groups = {}
    if use_grouping:
        for f in all_files:
            prefix = f[:2]
            groups.setdefault(prefix, []).append(f)
    else:
        groups[''] = all_files

    # 4. 执行合并逻辑
    for prefix, files in groups.items():
        if not files:
            continue

        if use_grouping:
            output_filename = f"{prefix}{folder_name}merge.xls"
        else:
            output_filename = f"{folder_name}_merge.xls"

        output_path = os.path.join(folder_path, output_filename)

        # 排除输出文件自身
        if output_filename in files:
            files.remove(output_filename)
            if not files: continue

        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception as e:
                print(f"无法删除已存在的输出文件 {output_filename}: {e}")
                continue

        print(f"\n>>> 正在生成 (按升序): {output_filename}")
        for idx, f in enumerate(files):
            print(f"    - 正在处理: {f}")

        # --- 开始合并流程 ---
        # A. 使用第一个文件作为母版
        template_file = os.path.join(folder_path, files[0])
        rb = xlrd.open_workbook(template_file, formatting_info=True)  # formatting_info 仅支持 xls

        # 查找目标 sheet 索引
        sheet_names = rb.sheet_names()
        if target_sheet not in sheet_names:
            print(f"错误：在文件 {files[0]} 中找不到名为 '{target_sheet}' 的工作表")
            continue

        target_sheet_index = sheet_names.index(target_sheet)
        wb_target = copy(rb)  # 复制为可写对象
        ws_target = wb_target.get_sheet(target_sheet_index)

        # 获取母版当前已有的行数
        current_row_idx = rb.sheet_by_index(target_sheet_index).nrows

        # B. 从第二个文件开始追加
        for i in range(1, len(files)):
            source_file_path = os.path.join(folder_path, files[i])
            rb_source = xlrd.open_workbook(source_file_path)

            if target_sheet not in rb_source.sheet_names():
                continue

            ws_source = rb_source.sheet_by_name(target_sheet)

            # 遍历源文件的行（跳过表头）
            for r_idx in range(header_rows, ws_source.nrows):
                row_values = ws_source.row_values(r_idx)

                # 检查是否整行全空
                if all(str(cell).strip() == "" for cell in row_values):
                    continue

                # 写入目标文件
                for col_idx, value in enumerate(row_values):
                    ws_target.write(current_row_idx, col_idx, value)
                current_row_idx += 1

        # C. 保存结果
        wb_target.save(output_path)
        print(f"--- {output_filename} 合并完成！---")

        # 5. 自动打开功能
        if os.name == 'nt':
            wps_path = r'F:\APP\wps\WPS Office\ksolaunch.exe'
            if os.path.exists(wps_path):
                subprocess.Popen([wps_path, output_path])
                print(f"已调用WPS打开: {output_filename}")
            else:
                print(f"警告：找不到WPS路径 {wps_path}，无法自动打开。")


# ==========================================
#               参数设置区域
# ==========================================
if __name__ == "__main__":
    # folder_name = "order"
    # 或者手动指定
    folder_name = "order"

    MY_FOLDER = os.path.join(r"C:\Users\Administrator\Desktop", folder_name)

    if not os.path.exists(MY_FOLDER):
        print(f"路径不存在: {MY_FOLDER}")
    else:
        merge_xls(
            folder_path=MY_FOLDER,
            header_rows=1,  # 表头占 1 行，从第 2 行开始复制
            target_sheet="账号业务统计",  # 确保你的 xls 里的 sheet 名正确
            use_grouping=True,
            exclude_merge=True
        )