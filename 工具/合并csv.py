import os
import pandas as pd

# 文件夹路径
FOLDER_PATH = r"F:\Super Browser\Super Browser\youleilei34@outlook.com\DGD广告"

# 输出文件名
OUTPUT_FILE = "合并结果.csv"
OUTPUT_PATH = os.path.join(FOLDER_PATH, OUTPUT_FILE)


def get_header_rows():
    """
    获取标题行数量，默认 1
    """
    value = input("请输入标题行数量，直接回车默认 1：").strip()

    if value == "":
        return 1

    try:
        rows = int(value)
        if rows < 0:
            print("标题行数量不能小于 0，已使用默认值 1")
            return 1
        return rows
    except ValueError:
        print("输入无效，已使用默认值 1")
        return 1


def read_csv_safely(file_path):
    """
    尝试用常见编码读取 CSV
    """
    encodings = ["utf-8-sig", "utf-8", "gbk", "gb18030"]

    for enc in encodings:
        try:
            return pd.read_csv(
                file_path,
                header=None,
                dtype=str,
                encoding=enc
            )
        except UnicodeDecodeError:
            continue

    # 如果常见编码都失败，最后用 gb18030 忽略错误
    return pd.read_csv(
        file_path,
        header=None,
        dtype=str,
        encoding="gb18030",
        errors="ignore"
    )


def merge_csv_files(folder_path, output_path, header_rows=1):
    all_data = []
    header_saved = False

    csv_files = [
        f for f in os.listdir(folder_path)
        if f.lower().endswith(".csv")
        and not f.startswith("~$")
        and f != os.path.basename(output_path)
    ]

    if not csv_files:
        print("文件夹内没有找到 CSV 文件")
        return

    for file_name in csv_files:
        file_path = os.path.join(folder_path, file_name)
        print(f"正在处理文件：{file_name}")

        try:
            df = read_csv_safely(file_path)

            # 删除完全空白行
            df = df.dropna(how="all")

            if df.empty:
                continue

            if header_rows > 0:
                if not header_saved:
                    # 第一个 CSV 保留标题行
                    all_data.append(df)
                    header_saved = True
                else:
                    # 后面的 CSV 跳过标题行
                    all_data.append(df.iloc[header_rows:])
            else:
                # 标题行数量为 0，则全部合并
                all_data.append(df)

        except Exception as e:
            print(f"处理文件失败：{file_name}")
            print(f"错误原因：{e}")

    if not all_data:
        print("没有可合并的数据")
        return

    merged_df = pd.concat(all_data, ignore_index=True)

    merged_df.to_csv(
        output_path,
        index=False,
        header=False,
        encoding="utf-8-sig"
    )

    print()
    print("合并完成！")
    print(f"文件已保存到：{output_path}")


if __name__ == "__main__":
    header_rows = get_header_rows()
    merge_csv_files(FOLDER_PATH, OUTPUT_PATH, header_rows)