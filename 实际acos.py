import csv


def process_csv(file_path):
    try:
        # 打开并读取 CSV 文件
        with open(file_path, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)

            # 读取标题行
            headers = next(reader)

            # 获取所需列的索引
            try:
                ad_activity_index = headers.index("广告活动")  # 找到“广告活动”列
                order_index = headers.index("订单")  # 找到“订单”列
                sales_index = headers.index("销售额(USD)")  # 找到“销售额(USD)”列
            except ValueError as e:
                print(f"错误：标题行中找不到所需列 - {e}")
                return

            # 初始化结果变量
            total_ad_order_sum = 0
            total_sales_sum = 0

            # 遍历每一行数据
            for row in reader:
                try:
                    # 提取“广告活动”列的内容，取最后4个字符并转化为浮点数
                    ad_activity_value = float(row[ad_activity_index][-4:])

                    # 提取“订单”列，将其转化为浮点数
                    order_value = float(row[order_index])

                    # 计算“广告活动 * 订单”，并累加
                    total_ad_order_sum += ad_activity_value * order_value

                    # 提取“销售额(USD)”列，将其转化为浮点数并累加
                    sales_value = float(row[sales_index])
                    total_sales_sum += sales_value

                except ValueError:
                    # 如果某行数据不能正确处理，跳过
                    print(f"警告：无法处理行数据 {row}")
                    continue

            # 计算并四舍五入结果
            total_ad_order_sum = round(total_ad_order_sum, 2)
            total_sales_sum = round(total_sales_sum, 2)
            total_output = round(total_sales_sum + total_ad_order_sum, 2)

            # 输出结果
            print(f"运费总金额为：{total_ad_order_sum}")
            print(f"销售额总金额为：{total_sales_sum}")
            print(f"广告产出为：{total_output}")

    except FileNotFoundError:
        print(f"错误：未找到文件 {file_path}")
    except Exception as e:
        print(f"发生错误：{e}")


# 调用函数，传入 CSV 文件路径
file_path = r"C:\Users\123\Desktop\250107.csv"  # 替换为你的 CSV 文件路径
process_csv(file_path)