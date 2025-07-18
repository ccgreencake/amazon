import openpyxl
from openpyxl import Workbook
import openpyxl
from openpyxl.chart import BarChart, Reference



def add_charts_optimized(output_file):
    # 加载工作簿
    wb = openpyxl.load_workbook(output_file)
    Sheet1 = wb['Sheet1']

    # 清理并准备图表工作表
    if 'Charts' in wb.sheetnames:
        charts_sheet = wb['Charts']
        # 清空现有内容但保留格式
        charts_sheet.delete_rows(1, charts_sheet.max_row)
        charts_sheet.delete_cols(1, charts_sheet.max_column)
    else:
        charts_sheet = wb.create_sheet("Charts")

    # 删除其他不需要的工作表（保留Sheet1和Charts）
    for sheet_name in wb.sheetnames[:]:
        if sheet_name not in ['Sheet1', 'Charts']:
            del wb[sheet_name]

    # 图表参数配置
    chart_width = 20
    chart_height = 10
    vertical_gap = 10  # 图表之间的行间距
    current_row = 1

    # 获取分类标签（B1-最后一列的标签）
    categories = Reference(Sheet1,
                           min_col=2, max_col=Sheet1.max_column,
                           min_row=1, max_row=1)

    # 遍历每个数据行
    for data_row in range(2, Sheet1.max_row + 1):
        # 跳过空值行
        if not Sheet1.cell(data_row, 1).value:
            continue

        # 创建图表对象
        chart = BarChart()
        chart.type = "col"
        chart.style = 4  # 更简洁的样式
        chart.title = Sheet1.cell(data_row, 1).value
        chart.width = chart_width
        chart.height = chart_height

        # 设置数据范围
        data = Reference(Sheet1,
                         min_col=2, max_col=Sheet1.max_column,
                         min_row=data_row, max_row=data_row)

        # 添加数据系列
        series = openpyxl.chart.Series(data, title="数值")
        chart.series.append(series)

        # 设置分类轴
        chart.set_categories(categories)

        # 配置坐标轴
        chart.x_axis.title = "分类"
        chart.y_axis.title = "数值"
        chart.y_axis.scaling.min = 0

        # 定位图表位置
        anchor = f"A{current_row}"
        charts_sheet.add_chart(chart, anchor)
        name = Sheet1.cell(data_row, 1).value  # 获取名称
        charts_sheet[f"L{current_row}"] = '@'+name  # 将名称插入到L列
        current_row += chart_height + vertical_gap

    # 保存前再次清理空工作表
    wb.save(output_file)
    print(f"优化后的图表已保存到 {output_file}")


output_file = r"C:\Users\123\Downloads\processed (6).xlsx"
add_charts_optimized(output_file)
