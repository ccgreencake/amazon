import random
import pandas as pd
import os
import numpy as np
import utils
import pinjie

# 配置常量
SOURCE_FILE_2_PATH = r'C:\Users\Administrator\Desktop\new 词表.xlsx'
IMG_IP = "192.3.95.71"


def title_case(s):
    """将字符串转换为标题形式，并处理特定介词和小写情况"""
    s = str(s).title()
    replacements = {
        'And': 'and', 'With': 'with', 'For': 'for',
        "'S": "'s", "’S": "'s"
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    return s


def get_account_config(account):
    """根据账号获取配置信息"""
    config = {
        "LIANG": {"prefix": "A", "brand": "Bakgeerle", "domain": "lian.yant88.xyz"},
        "YOULE": {"prefix": "B", "brand": "Bakgeerle", "domain": "youl.yant88.xyz"},
        "QMTIY": {"prefix": "B", "brand": "Generic", "domain": "siya.yant88.xyz"}
    }
    return config.get(account, {"prefix": "A", "brand": "Generic", "domain": IMG_IP})


def process_images(df_source, domain, zuhe, length):
    """处理图片链接替换逻辑"""
    # 定义图片列索引映射
    img_cols_map = {
        'main_img': 11, 'main_img_main': 9, 'img_1': 13,
        'img_2': 9, 'img_3': 15, 'img_4': 17,
        'img_5': 19, 'img_6': 21, 'img_7': 23,
        'final_img': 25  # 实际上代码里img_6用了21, final_img用了25(尺码表)
    }

    imgs = {}

    # 提取并替换域名
    def replace_url(urls, is_single=False):
        if is_single:
            return str(urls).replace(IMG_IP, domain) if pd.notna(urls) else urls
        return [str(url).replace(IMG_IP, domain) if pd.notna(url) else url for url in urls]

    for key, col_idx in img_cols_map.items():
        if key == 'main_img_main':
            imgs[key] = replace_url(df_source.iloc[0, col_idx], is_single=True)
        else:
            raw_data = df_source.iloc[0:, col_idx].tolist()
            imgs[key] = replace_url(raw_data)

    if zuhe == 1:
        imgs['img_zuhe'] = replace_url(df_source.iloc[0:, 13].tolist())

    return imgs


def amazon(source_file, CiBiao, account, TG_VP, price, shipping, leixing, style, fangfa, color_num, zuhe, zibiao,
           item_type, manufacturer):
    # 1. 基础路径和文件处理
    df_source = pd.read_excel(source_file)
    folder_name = os.path.basename(os.path.dirname(source_file))
    file_name = os.path.basename(source_file)
    new_string = os.path.splitext(file_name)[0].replace("product_", "")

    df_source2 = pd.read_excel(SOURCE_FILE_2_PATH, sheet_name=CiBiao)

    # 2. 获取配置
    acct_conf = get_account_config(account)
    brand = acct_conf['brand']
    color_prefix = acct_conf['prefix']

    # 3. 价格计算
    list_price = float(price) + 10 if float(price) <= 15 else price

    # 4. 提取词表数据 (Bullet Points & Keywords)
    # word = df_source2.iloc[0:, 0].dropna() # 代码中未使用，暂时注释
    if zibiao == 0:
        zi = df_source2.iloc[0:, 2].dropna()
    point5 = pinjie.pinjie2(5, CiBiao, account)

    # 5. SKU 生成
    sku_old = df_source.iloc[0:, 0]
    sku_prefix = f"{TG_VP}LDB{account}-{folder_name}-{new_string}-"
    sku = [f"{sku_prefix}{s}" for s in sku_old]
    length = len(sku)

    # 6. Title 处理
    title_raw = df_source.iloc[0:, 2]
    if TG_VP == 'P':
        title_raw[:] = title_raw.iloc[color_num]
    base_title_str = title_raw[0]
    rise = utils.rise(base_title_str)

    # 7. Color 生成 (核心逻辑保持不变)
    colors = df_source.iloc[0:, 4].tolist()  # 转为list处理
    size = df_source.iloc[0:, 5]
    size_count = size.nunique()
    new_color = []

    # 这里的逻辑比较复杂，保留原逻辑
    n = 0
    count = 1
    flag = 1
    digits = [i + 1 for i in range(len(colors) // size_count + 1)]
    random.shuffle(digits)

    for i in range(len(colors)):
        if i != 0:
            if colors[i] != colors[i - 1]:
                n += 1
                if i != 1: flag = 0
            elif flag == 1:
                count += 1

        if TG_VP == 'P':
            prefix = f'P{color_num}{str(digits[n]).zfill(2)}- '
        elif color_num != 0:
            prefix = f'C{color_num}{str(n).zfill(2)}- '
        else:
            prefix = f'{color_prefix}{str(n).zfill(3)}- '
        new_color.append(prefix + str(colors[i]))

    # 8. Size Chart & Points处理
    size_chart = df_source.iloc[1, 26]
    p5 = df_source.iloc[0:, 27]
    if TG_VP == 'P':
        mask = p5.notna()
        p5[mask] = np.random.permutation(p5[mask].values)

    # 清理 p5 中的中文符号
    p5 = [str(x).split('【')[-1] if '【' in str(x) else str(x) for x in p5]

    # 9. 图片处理 (使用封装函数)
    imgs = process_images(df_source, acct_conf['domain'], zuhe, length)

    # 10. 构建新 DataFrame
    df_new = pd.DataFrame()

    # --- 基础列 ---
    feed_product = 'onepieceoutfit' if leixing in ['jumpsuit', 'romper'] else leixing
    df_new['Column1'] = [feed_product] * length
    df_new['Column2'] = sku
    df_new['brand'] = [brand] * length
    df_new['update'] = ['Update'] * length

    # --- 标题处理 ---
    # colors[0] = np.nan # 原逻辑修改了原数据，这里如果不影响后续，保留
    if zibiao == 1:
        new_title = title_raw
    else:
        # 注意：这里传入的 colors 列表第一个元素在原代码中被设为了 NaN
        colors_for_zibiao = df_source.iloc[0:, 4].tolist()
        colors_for_zibiao[0] = np.nan
        new_title = pinjie.zibiao(length, CiBiao, title_raw, colors_for_zibiao)

    df_new['item_name'] = [title_case(t) for t in new_title]

    # --- Item Type 逻辑 ---
    final_item_type = np.nan
    if TG_VP == 'P':
        final_item_type = [item_type] * length
    else:
        cb_lower = str(CiBiao).lower()
        if any(x in cb_lower for x in ['weiku', 'sweatpant', 'jogger']):
            vals = ['athletic-sweatpants', 'casual-pants', 'athletic-dance-pants', 'athletic-maternity-pants',
                    'athletic-pants', 'athletic-cricket-pants', 'athletic-track-pants']
            final_item_type = [vals[i % len(vals)] for i in range(length)]
        elif 'tactical pants' in cb_lower or 'cargo pants' in cb_lower:
            vals = ['hiking-pants', 'casual-pants', 'military-pants']
            final_item_type = [vals[i % len(vals)] for i in range(length)]
        elif 'track suit' in cb_lower:
            vals = ['athletic-tracksuits', 'athletic-sweatsuits']
            final_item_type = [vals[i % len(vals)] for i in range(length)]
        else:
            final_item_type = np.nan

    df_new['item_type'] = final_item_type

    # --- Sport / Activity ---
    if 'cargo' in CiBiao or 'tactical' in CiBiao:
        s1, s2 = 'Tactical & Military', 'Hiking'
    elif 'tennis' in CiBiao or 'golf' in CiBiao:
        s1, s2 = 'Tennis', 'Golf'
    else:
        s1, s2 = 'Outdoor Lifestyle', 'Walking'
    df_new['sport1'] = [s1] * length
    df_new['sport2'] = [s2] * length

    # --- Manufacturer ---
    if TG_VP == 'P':
        df_new['manufacturer'] = [manufacturer] * length
    else:
        df_new['manufacturer'] = [brand] if account in ["YOULE", "LIANG"] else ['Generic'] * length

    df_new['style'] = [style] * length
    # title_small 替换为 .lower()
    df_new['key'] = pinjie.pinjie1(length, CiBiao, [t.lower() for t in title_raw], fangfa)
    df_new['weave_type'] = ['Knit'] * length

    # --- Points (Bullet Points) ---
    if TG_VP == "P":
        for i in range(5):
            df_new[f'point{i}'] = [point5[i]] * length
    else:
        # 这里原来的逻辑是取p5列的前5个非空值，原代码p5[0]..p5[4]
        for i in range(5):
            # 确保p5长度足够，防止越界
            val = p5[i] if i < len(p5) else ""
            df_new[f'p{i + 1}'] = [val] * length

    df_new['color'] = new_color
    df_new['color_map'] = df_source.iloc[0:, 4].tolist()  # 原始颜色

    # --- 常量填充 ---
    common_cols = {
        'belt_style': 'Medium', 'control_type': 'Medium', 'front_style': 'Flat Front',
        'country_as_labeled': 'China', 'pattern_name': 'solid', 'pattern_type': 'solid',
        'toe_style': 'Footless', 'material_type': 'Polyester', 'lifestyle': 'Casual',
        'fit_type': 'Regular', 'rise_style': rise, 'is_autographed': 'No',
        'item_weight_unit_of_measure': 'GR', 'item_weight': '150', 'length': 'regular',
        'care_instructions': 'Machine Wash', 'import_designation': 'Imported',
        'closure_type': 'Elastic', 'inner_material_type': 'Polyester',
        'outer_material_type1': 'Polyester Blend', 'outer_material_type2': 'Cotton Blend',
        'outer_material_type3': 'Polyester', 'outer_material_type4': 'Polyester',
        'outer_material_type5': 'Polyester', 'part_number': 'Bakgeerle',
        'collection_name': 'All', 'shaft_style_type': '', 'currency': 'USD',
        'condition_type': 'New', 'shipping': shipping, 'package_level': 'unit',
        'package_contains_quantity': '1', 'item_length_unit_of_measure': 'CM',
        'item_length': '25', 'item_width': '20', 'item_height': '2.0',
        'item_width_unit_of_measure': 'CM', 'item_height_unit_of_measure': 'CM',
        'package_height': '2', 'package_width': '20', 'package_length': '25',
        'package_length_unit_of_measure': 'CM', 'package_weight': '150',
        'package_weight_unit_of_measure': 'GR', 'package_height_unit_of_measure': 'CM',
        'package_width_unit_of_measure': 'CM'
    }
    for col, val in common_cols.items():
        df_new[col] = [val] * length

    # Occasion loop
    for i in range(1, 6):
        df_new[f'occasion_type{i}'] = ['Casual'] * length

    # --- 逻辑分支填充 ---
    # Pocket
    if 'cargo' in CiBiao or 'tactical' in CiBiao:
        vals = ['Cargo', 'Slit Pockets', 'Carpenter', 'Utility Pocket']
        df_new['pocket_description'] = [vals[i % len(vals)] for i in range(length)]
    else:
        df_new['pocket_description'] = ['Slit Pocket'] * length

    # Lifecycle
    df_new['lifecycle_supply_type'] = ['Year Round Replenishable' if i % 2 == 0 else 'Fashion' for i in range(length)]

    # Leg Style
    if "Flare" in base_title_str:
        leg_s = 'Flared'
    elif "Cropped" in base_title_str:
        leg_s = 'Cropped'
    else:
        leg_s = 'Wide'
    df_new['leg_style'] = [leg_s] * length

    # Water Resistance
    if 'tactical' in CiBiao:
        vals = ['water_resistant', 'Waterproof']
        df_new['water_resistance_level'] = [vals[i % len(vals)] for i in range(length)]
    else:
        df_new['water_resistance_level'] = ['not_water_resistant'] * length

    # Department / Gender
    dept, gender = ('Men', 'Male') if 'men' in CiBiao else ('Women', 'Female')
    df_new['department'] = [dept] * length
    df_new['gender'] = [gender] * length

    # Fabric / Fur
    # 注意：这里使用point5列表(从pinjie2获取的)
    df_new['fur_description'] = [point5[2]] * length
    if feed_product == 'pants':
        df_new['fabric_type'] = [point5[3]] * length  # 原代码此处用p5[3]有时混淆，逻辑上point5更稳，保持原逻辑用p5
        # 修正：原代码在 feed_product == 'pants' 用了 p5[3]，在 else 用了固定值
        # 这里需要注意 p5 在上面被处理过。假设 pinjie.pinjie2 返回的 point5 和 Excel读的 p5 逻辑分离
        # 按照原代码逻辑：df_new['fabric_type'] = [p5[3]]*length
        df_new['fabric_type'] = [p5[3] if len(p5) > 3 else ""] * length
    else:
        df_new['fabric_type'] = [
                                    '55%linen, 45%cotton' if 'Linen' in base_title_str else '71%Polyester,18%Cotton,11%Spandex'] * length

    # Description
    if TG_VP in ["C", "P"]:
        product_description = utils.description("../html.txt", size_chart, p5)
    else:
        product_description = utils.description("../html.txt", size_chart, p5)  # 逻辑一致
    df_new['product_description'] = product_description
    if len(product_description) > 2000:
        print("大描述超过2000字符!")

    # Model Name
    model_val = 'Bakgeerle' if account in ["YOULE", "LIANG"] else 'Generic'
    df_new['model'] = [model_val] * length
    df_new['model_name'] = [model_val] * length

    # Price / Quantity / Age
    df_new['price'] = [price] * length
    df_new['quantity'] = ['1000'] * length
    df_new['age'] = ['Adult'] * length

    # --- Size System Logic ---
    df_new['size'] = size
    # 初始化 size 相关固定列
    for col in ['size_sys', 'Alpha', 'body_type', 'height_type']:
        df_new[col] = np.nan  # 默认空，根据类型填充

    # 根据类型填充空白列 (bai0...baiN)
    bai_range = (0, 0)  # start, count
    if leixing in ['pants', 'jumpsuit', 'shorts', 'romper', 'overalls', 'dress', 'nightgownnightshirt', 'sweatshirt',
                   'sweater', 'coat', 'suit', 'tracksuit', 'skirts', 'skirt', 'shirt']:
        df_new['size_sys'] = ['US'] * length
        df_new['Alpha'] = ['Alpha'] * length
        df_new['body_type'] = ['Regular'] * length
        df_new['height_type'] = ['Regular'] * length

    # 填充空白列逻辑 (原代码逻辑)
    if leixing in ['pants', 'jumpsuit', 'shorts', 'romper', 'overalls']:
        for i in range(15): df_new[f'bai{i}'] = np.nan
        df_new['waist_style'] = ['Medium Waist'] * length
    elif leixing in ['dress', 'nightgownnightshirt', 'sweatshirt', 'sweater', 'coat', 'suit', 'tracksuit']:
        for i in range(5): df_new[f'bai{i}'] = np.nan
        for i in range(10): df_new[f'bai{i + 5}'] = np.nan
    elif leixing in ['skirts', 'skirt']:
        for i in range(10): df_new[f'bai{i}'] = np.nan
        for i in range(5): df_new[f'bai{i + 10}'] = np.nan
    elif leixing == 'shirt':
        for i in range(15): df_new[f'bai{i}'] = np.nan

    # --- Images ---
    # 如果 zuhe=1 或 TG_VP 为 P/C，全部使用主图列表，否则根据 count 混合
    df_new['main_img'] = imgs['main_img']
    if not (zuhe == 1 or TG_VP in ["P", "C"]):
        # 使用 loc 切片赋值
        limit = min(int(count) + 1, len(df_new))  # 防止越界
        df_new.loc[:limit, 'main_img'] = imgs['main_img_main']

    df_new['img_1'] = imgs['img_1']
    df_new['img_2'] = imgs['img_2']
    df_new['img_3'] = imgs['img_3']
    df_new['img_4'] = imgs['img_4']
    df_new['img_5'] = imgs['img_5']
    df_new['img_6'] = imgs['final_img']  # 对应原代码 final_img
    df_new['img_7'] = imgs['img_7']
    df_new['final_img'] = imgs['final_img']  # 对应原代码 img_6
    df_new['swatch_img'] = imgs['main_img']

    # --- 变体关系 (Parent/Child) ---
    df_new['parent_child'] = ['Child'] * length
    df_new['parent_sku'] = [sku[0]] * length
    df_new['relationship'] = ['Variation'] * length
    df_new['variation_theme'] = ['color-size'] * length

    # --- Size Mapping ---
    size_map = [
        str(s).replace("2X", "XX").replace("3X", "XXX").replace("4X", "XXXX").replace("5X", "XXXXX").replace("6X",
                                                                                                             "XXXXXX")
        for s in size]
    df_new['size_map'] = size_map
    df_new['size_name'] = size_map
    df_new['list_price'] = [list_price] * length

    # --- 清理第一行 (Parent Row) ---
    cols_to_clean = [
        'color', 'color_map', 'leg_style', 'is_autographed', 'department',
        'price', 'quantity', 'age', 'Alpha', 'size_sys', 'body_type', 'height_type',
        'main_img', 'img_1', 'img_2', 'img_3', 'img_4', 'img_5', 'img_6', 'img_7',
        'final_img', 'swatch_img', 'fit_type', 'rise_style', 'size_map', 'size_name',
        'list_price', 'currency', 'condition_type', 'shipping', 'parent_sku', 'relationship'
    ]
    df_new.loc[0, cols_to_clean] = np.nan

    # 设置 Parent 行特有值
    df_new.loc[0, 'parent_child'] = 'Parent'

    # 11. 埋词扩充逻辑 (如果需要启用，可以恢复，原代码主要逻辑在计算 num 后 break 了)
    if zibiao == 0:
        a = 17
        num = length * a
        # 简单的扩充逻辑
        if num > len(zi):
            repeat_times = (num // len(zi)) + 2
            zi = pd.concat([zi] * repeat_times, ignore_index=True)
        zi = zi.iloc[:num]
        # 原代码这里并没有把埋词真正赋值给 df_new 的列，只是准备了数据。保持原样。

    # 12. 保存文件
    output_filename = f"{os.path.splitext(file_name)[0]}.xlsx"
    df_new.to_excel(output_filename, index=False)

    return os.path.join(r'/', output_filename)