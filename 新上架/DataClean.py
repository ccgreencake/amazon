import pandas as pd
import datetime
import os


class SKUGenerator:
    """
    专门负责亚马逊 SKU 生成、颜色处理与原始数据提取的类
    """

    def __init__(self, base_dir=r'F:\product'):
        self.base_dir = base_dir

    def get_dynamic_info(self, file_index):
        date_str = datetime.datetime.now().strftime('%y%m%d')
        formatted_index = str(file_index).zfill(2)
        folder_path = os.path.join(self.base_dir, date_str)
        return date_str, formatted_index, folder_path

    def generate_custom_colors(self, df, color_column='颜色'):
        """
        给颜色添加前缀：识别唯一颜色并分配 C101, C102...
        :param df: 传入的 DataFrame
        :param color_column: 原始颜色列的名称，默认为 '颜色'
        :return: 带有 '自定义颜色' 列的 DataFrame
        """
        if color_column not in df.columns:
            print(f"⚠️ 警告：未在文件中找到列名 '{color_column}'，跳过颜色处理。")
            return df

        # 1. 获取唯一的颜色列表（按出现顺序排序）
        # unique() 会保留第一次出现的顺序
        unique_colors = df[color_column].unique()

        # 2. 构建颜色映射字典
        # 例如：{'Blue': 'C101- Blue', 'Grey': 'C102- Gray'}
        color_mapping = {}
        for i, color in enumerate(unique_colors):
            code = 101 + i  # 从 101 开始计数

            # 特殊处理：根据你的示例，Grey 变成了 Gray
            display_name = str(color).strip()
            if display_name.lower() == 'grey':
                display_name = 'Gray'

            new_value = f"C{code}- {display_name}"
            color_mapping[color] = new_value

        # 3. 将映射应用到新列 '自定义颜色'
        df['自定义颜色'] = df[color_column].map(color_mapping)

        print(f"✅ 颜色处理完成：已生成 {len(unique_colors)} 种自定义颜色。")
        return df

    def run_sku_process(self, file_index, char_prefix='T', brand='YOULE'):
        """
        执行核心流程
        """
        date_str, formatted_index, folder_path = self.get_dynamic_info(file_index)
        file_name = f"product_{formatted_index}.xls"
        full_file_path = os.path.join(folder_path, file_name)

        if not os.path.exists(full_file_path):
            raise FileNotFoundError(f"错误：未找到源文件 {full_file_path}")

        try:
            # 读取原始 Excel
            df = pd.read_excel(full_file_path)
        except Exception as e:
            raise Exception(f"读取 Excel 失败: {str(e)}")

        # --- 处理 SKU ---
        sku_prefix = f"{char_prefix}LDB{brand}-{date_str}-{formatted_index}"
        original_data = df.iloc[:, 0].astype(str)
        df['custom_sku'] = original_data.apply(lambda x: f"{sku_prefix}-{x}")

        # --- 处理颜色 (调用新方法) ---
        # 假设原始文件里有一列叫 '颜色'，如果没有，你可以根据实际列名修改
        df = self.generate_custom_colors(df, color_column='颜色')

        print(f"成功处理文件：{file_name}")
        return df


# ==========================================
# 测试代码块
# ==========================================
if __name__ == "__main__":
    processor = SKUGenerator()
    try:
        # 模拟运行
        final_data = processor.run_sku_process(file_index=1)

        # 查看结果中新增的两列
        print("\n--- 处理结果预览 ---")
        if '自定义颜色' in final_data.columns:
            print(final_data[['颜色', '自定义颜色', 'custom_sku']].head(15))
    except Exception as e:
        print(f"程序运行中出现问题：{e}")