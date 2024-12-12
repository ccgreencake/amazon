import pandas as pd
import os
import subprocess


def classify_cells_by_content(xlsx_file, categories, out_path):
    sheet_name = 'Sheet1'
    # 读取Excel文件
    df = pd.read_excel(xlsx_file, sheet_name=sheet_name)

    # 初始化结果DataFrame，并将分类名称设置为表头
    result_df = pd.DataFrame(columns=categories + ['其他颜色'])

    # 遍历首列，根据内容分类
    for index, cell in df.iloc[:, 0].items():
        cell_lower = str(cell).lower()
        # 检查每个分类，并将匹配的单元格值放入对应的列
        for category in categories:
            if category in cell_lower:
                # 如果该列已经有数据，则找到最后一个非空值的索引并加1
                if not result_df[category].isnull().all():
                    last_index = result_df[category].last_valid_index()
                    result_df.at[last_index + 1, category] = cell
                else:
                    # 如果该列没有数据，则直接添加到第一行
                    result_df.at[0, category] = cell
                break
        else:
            # 如果没有匹配的分类，则放入'Default'列
            if not result_df['其他颜色'].isnull().all():
                last_index = result_df['其他颜色'].last_valid_index()
                result_df.at[last_index + 1, '其他颜色'] = cell
            else:
                result_df.at[0, '其他颜色'] = cell

    # 删除完全为空的列
    result_df = result_df.apply(lambda x: x.dropna(how='all'), axis=0)

    # 保存结果到新的Excel文件
    result_df.to_excel(out_path, index=False)
    return result_df


# 使用示例
xlsx_file = rf'C:\Users\123\Desktop\工作簿1.xlsx'  # Excel文件路径
  # 工作表名称
# categories = ['white', 'black','blue','green','red','purple','burgundy','pink','yellow','orange','grey','gray','khaki','camo','brown','tan']  # 分类关键词列表
# categories = ['low rise',	'high waist',	'wide leg',	'straight leg',	'fleece',	'plus size',	'cargo',]  # 分类关键词列表
# categories = ['low rise','high waist','wide leg','drawstring','baggy','loose','slim','relax','water']  # 分类关键词列表
# categories = ['relax','regular','slim']  # 分类关键词列表
# categories = ['flannel','fleece','plaid','christmas','halloween','family','fluffy','fuzzy']
# categories = ['xs','6xl','5xl','4xl','3xl','2xl','xl']
# categories = ['6x','5x','4x','3x','2x']
# categories = ['oversized', 'long', 'flannel', 'fleece', 'cotton', 'scrub', 'wool', 'insulated', 'lightweight', 'thick', 'heavyweight', 'warm', 'thermal', 'high', 'low', 'pleated', 'flowy', 'plus size', 'large', 'petite', 'gym', 'athletic', 'casual', 'soft', 'comfy', 'cool', 'comfortable', 'relaxed fit', 'skinny', 'baggy', 'slim', 'loose']
# categories = ['wide leg', 'heavy', 'running', 'cargo', 'jogger', 'cold', 'pocket', 'christmas', 'fuzzy', 'sock', 'velvet', 'twill', 'jean', 'denim', 'winter', 'yoga', 'rise', 'plaid', 'big', 'small', 'tall']
# categories = ['christmas','lounge','leg','drawstring']
# categories = ['for men','mens',"men's","men",'man']
categories = ['graphic', 'plaid', 'christmas', 'waist', 'drawstring', 'low', 'high', 'relaxed fit', 'loose', 'slim',
              'baggy', 'skinny', 'casual', 'gym', 'lounge', 'athletic', 'thick', 'warm', 'thermal', 'heavyweight',
              'lightweight', 'fuzzy', 'insulated', 'scrub', 'flannel', 'cotton', 'lined', 'fleece', 'winter', 'wool',
              'small', 'long', 'oversized', 'tall', 'big', 'stack', 'flare', 'cuff', 'stretch', 'comfy', 'tapered',
              'fleece lined', 'rip', 'waterproof', 'elastic bottom', 'open bottom', 'wide leg', 'leg', 'chino', 'dress',
              'yoga', 'jean', 'running', 'cargo', 'tactical', 'work', 'cargos', 'military', 'jogger', 'velvet',
              'pocket', 'man', "men's", 'men', 'mens', 'for men']
categories = ['combat', 'utility', 'camping', 'security', 'hunting', 'carpenter', 'fishing', 'light', 'work', 'hiking',
              'cargo']
out_path = rf'C:\Users\123\Desktop\classified.xlsx'
# 调用函数并保存结果
classified_data = classify_cells_by_content(xlsx_file, categories, out_path)
if os.name == 'nt':  # Windows系统
    # 假设WPS的路径是 "C:\Program Files\WPS Office\et.exe"
    wps_path = r'C:\Users\123\AppData\Local\Kingsoft\WPS Office\ksolaunch.exe'
    # 使用WPS打开Excel文件
    subprocess.call([wps_path, out_path])



