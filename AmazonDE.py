import random

import pandas as pd
import os
import numpy as np
import utils
import pinjie


def title_case(s):
    # 将字符串转换为首字母大写的标题形式
    s = s.title()
    # 替换 'And', 'With', 'For' 为小写形式
    s = s.replace('And', 'and')
    s = s.replace('With', 'with')
    s = s.replace('For', 'for')
    s = s.replace("'S", "'s")
    s = s.replace("’S", "'s")
    return s
def title_small(s):
    s = s.lower()
    return s
def amazon(mubiao_file,source_file,CiBiao,account,TG_VP,price,shipping,leixing,style,fangfa,color_num,zuhe,zibiao,item_type,manufacturer):
    # 读取源表格文件
    # 替换为实际的源文件路径
    df_source = pd.read_excel(source_file)
    source_file2 = r'C:\Users\Administrator\Desktop\词表DE.xlsx'  # 替换为实际的源文件路径
    # 首先获取文件的目录路径
    directory_path = os.path.dirname(source_file)
    # 然后获取目录路径的最后一部分，即 '240607'
    folder_name = os.path.basename(directory_path)

    if (float(price)<=15):
        list_price = float(price)+10
    else:
        list_price = price



    if (account == "SANSK"):
        color_prefix = "A"
        brand = "Generisch"
    elif (account == "YOULE"):
        color_prefix = "B"
        brand = "Bakgeerle"
    elif(account == "zz"):
        color_prefix = "B"
        brand = "Generic"
        



    df_source2 = pd.read_excel(source_file2, sheet_name=CiBiao)

    #提取埋词
    word = df_source2.iloc[0:, 0]
    word = word.dropna()
    # 提取子标词
    if zibiao == 0:
        zi = df_source2.iloc[0:, 2]
        zi = zi.dropna()
    # point5 = df_source2.iloc[0:, 2]
    point5 = pinjie.pinjie2(5,CiBiao,account)



    # 获取表格文件的名称
    file_name = os.path.basename(source_file)
    table_name = os.path.splitext(file_name)[0]

    prefix = "product_"
    new_string = table_name[len(prefix):]

    # 提取sku
    sku_old = df_source.iloc[0:, 0]

    sku = []
    prefix = TG_VP+'LDB'+account+'-'+folder_name+'-'+new_string+'-'

    for string in sku_old:
        new_string = prefix + str(string)
        sku.append(new_string)

    # 提取title
    title = df_source.iloc[0:, 2]
    if TG_VP == 'P':
        title[:] = title.iloc[color_num]
    string = title[0]
    rise = utils.riseDE(string)


    # 提取color
    colors = df_source.iloc[0:, 4]
    n=0
    size = df_source.iloc[0:, 5]
    size_count = size.nunique()
    new_color = []
    count  = 1
    flag = 1
    digits = [i + 1 for i in range(len(colors) // size_count+1)]
    random.shuffle(digits)
    for i in range(len(colors)):
        if i != 0:
            if colors[i] != colors[(i-1)]:
                n = n+1
                if i != 1:
                    flag = 0
            elif flag ==1:
                count = count+1
        if TG_VP == 'P':
            prefix = 'P'+str(color_num)+str(digits[n]).zfill(2) + "- "
        elif color_num != 0:
            prefix = 'C'+str(color_num)+str(n).zfill(2)+"- "
        else:
            prefix = color_prefix + str(n).zfill(3) + "- "
        color = prefix + str(colors[i])
        new_color.append(color)
    # 提取size

    size_chart = df_source.iloc[1, 26]
    p5 = df_source.iloc[0:, 27]
    if TG_VP == 'P':
        mask = p5.notna()  # 标记非空值的位置
        non_null_indices = p5[mask].index  # 获取非空值的索引

        # 步骤3：提取非空值并打乱顺序
        non_null_values = p5[mask].values
        shuffled_values = np.random.permutation(non_null_values)  # 打乱

        # 步骤4：将打乱后的值赋回 p5 的非空位置
        p5[non_null_indices] = shuffled_values
    print('count:',count)
    # new_color = colors
    # 提取价格
    # price_str = df_source.iloc[0:, 7]
    # price = []
    # list_price=[]

    # for string in price_str:
    #     start_index = string.find("*US灵境：") + len("*US灵境：")
    #     end_index = string.find(" USD  *US顺丰")
    #     new_string = string[start_index:end_index]
    #     rounded_float = round(float(new_string), 1) - 2.01
    #     rounded_float1 = rounded_float+6
    #     formatted_string = "{:.2f}".format(rounded_float)
    #     formatted_string1 = "{:.2f}".format(rounded_float1)
    #     price.append(formatted_string)
    #     list_price.append(formatted_string1)
    # df_source.replace('*192.3.95.71*', '192.3.95.71', inplace=True)
    #提取图片
    main_img_1 = df_source.iloc[0:,11]
    main_img_main = df_source.iloc[0, 9]
    img_1_1 = df_source.iloc[0:,13]
    img_2_1 = df_source.iloc[0:,9]
    img_3_1 = df_source.iloc[0:,15]
    img_4_1 = df_source.iloc[0:,17]
    img_5_1 = df_source.iloc[0:,19]
    img_6_1 = df_source.iloc[0:,21]
    img_7_1 = df_source.iloc[0:,23]
    if zuhe == 1:
        img_zuhe_1 = df_source.iloc[0:,13]
    #     最后一张是尺码表
    # final_img_1 = df_source.iloc[0:,13]
    # 第二张是尺码表
    final_img_1 = df_source.iloc[0:,25]
    main_img = main_img_1.tolist()
    final_img = final_img_1.tolist()
    img_1 = img_1_1.tolist()
    img_2 = img_2_1.tolist()
    img_3 = img_3_1.tolist()
    img_4 = img_4_1.tolist()
    img_5 = img_5_1.tolist()
    img_6 = img_6_1.tolist()
    img_7 = img_7_1.tolist()
    if zuhe == 1:
        img_zuhe = img_zuhe_1.tolist()
    if (account == "zz"):
        main_img = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in main_img]
        main_img_main = main_img_main.replace("192.3.95.71", "192.3.95.71")
        final_img = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in final_img]
        img_1 = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in img_1]
        img_2 = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in img_2]
        img_3 = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in img_3]
        img_4 = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in img_4]
        img_5 = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in img_5]
        img_6 = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in img_6]
        img_7 = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in img_7]
        if zuhe == 1:
            img_zuhe = [url.replace("192.3.95.71", "192.3.95.71") if not pd.isna(url) else url for url in img_zuhe]
    elif account == "SANSK":
        main_img = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in main_img]
        main_img_main = main_img_main.replace("192.3.95.71", "youl.yant88.xyz")
        final_img = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in final_img]
        img_1 = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in img_1]
        img_2 = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in img_2]
        img_3 = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in img_3]
        img_4 = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in img_4]
        img_5 = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in img_5]
        img_6 = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in img_6]
        img_7 = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in img_7]
        if zuhe == 1:
            img_zuhe = [url.replace("192.3.95.71", "youl.yant88.xyz") if not pd.isna(url) else url for url in
                        img_zuhe]



    # 创建新的数据框架
    df_new = pd.DataFrame()

    length = len(sku)  # 填充的长度
    string = title[0]
    if leixing == 'jumpsuit' or leixing == 'romper':
        feed_product = 'onepieceoutfit'
    else:
        feed_product = leixing

    df_new['Column1'] = [feed_product] * length

    # 插入目标数据到第二列
    df_new['Column2'] = sku

    df_new['brand'] = [brand] * length
    df_new['update'] = ['Aktualisierung'] * length
    colors[0] = np.nan
    # new_title = [str(f_title) + ' ' + str(color) if i > 0 else f_title for i, (f_title, color) in enumerate(zip(title, colors))]
    small_title = [title_small(i) for i in title]
    # small_title.pop(0)
    if zibiao == 1:
        new_title = title
    elif zibiao == 0:
        new_title = pinjie.zibiao(length,CiBiao,title,colors)
    new_title = [title_case(i) for i in new_title]

    # for i in new_title:
    #     if len(i) > 200:
    #         print("第",i,"个标题长度超过200字符!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    #         break

    df_new['item_name'] = new_title

    # item_type = utils.item_type(string, feed_product)
    # df_new['item_type'] = [item_type] * length
    if TG_VP == 'P':
        df_new['item_type'] = [item_type] * length
    else:
        df_new['item_type'] = np.nan
    df_new['neck_type'] = np.nan
    df_new['sleeve_type'] = np.nan
    if 'cargo' in CiBiao or 'tactical' in CiBiao:
        df_new['sport1'] = ['Taktisch & Militär'] * length
        df_new['sport2'] = ['Wandern'] * length
    elif 'tennis' in CiBiao or 'golf' in CiBiao:

        df_new['sport1'] = ['Tennis'] * length
        df_new['sport2'] = ['Golf'] * length
    else:
        df_new['sport1'] = ['Outdoor-Lifestyle'] * length
        df_new['sport2'] = ['Gehen'] * length


    if TG_VP == 'P':
        df_new['manufacturer'] = [manufacturer] * length
    else:
        df_new['manufacturer'] = ['Generisch'] * length
    # df_new['model'] = np.nan
    # df_new['model_name'] = np.nan
    df_new['style'] = [style] * length
    df_new['key'] = pinjie.pinjie1(length,CiBiao,small_title,fangfa)
    # df_new['weave_type'] = ['Knit'] * length
    if TG_VP == "P":
        for i in range(5):
            name = 'point' + str(i)
            df_new[name] = [point5[i]] * length
    else:
        df_new['p1'] = [p5[0]] * length
        df_new['p2'] = [p5[1]] * length
        for i in range(3):
            name = 'point'+str(i)
            df_new[name] = [point5[i]] * length


    df_new['color'] = new_color
    df_new['color_map'] = colors
    df_new.loc[0, ['color', 'color_map']] = np.nan
    # df_new['belt_style']=['Medium']* length
    # df_new['control_type']=['Medium']* length
    df_new['front_style']=['Flache Vorderseite']* length
    df_new['country_as_labeled']=['China'] * length
    # df_new['pattern_name']=['solid']* length
    # df_new['pattern_type']=['solid']* length
    if 'cargo' in CiBiao or 'tactical' in CiBiao:
        values = ['Cargo']
        df_new['pocket_description'] = [values[i % len(values)] for i in range(length)]
    else:
        df_new['pocket_description']=['Slit Pocket']* length
    # df_new['toe_style']=['Footless'] * length
    # df_new['material_type']=['Polyester']* length
    df_new['lifestyle']=['Normal']* length
    df_new['lifecycle_supply_type'] = ['Year Round Replenishable' if i % 2 == 0 else 'Fashion' for i in range(length)]
    df_new['fit_type']=['Normal'] * length

    df_new['rise_style'] = [rise] * length
    if "Flare" in string:
        leg_style = 'Ausgestellt'
    elif "Cropped" in string:
        leg_style = 'Abgeschnitten'
    else:
        leg_style = 'Weit'
    df_new['leg_style'] = [leg_style] * length
    df_new.loc[0, 'leg_style'] = np.nan

    if 'cargo' in CiBiao or 'tactical' in CiBiao:
        values = ['water_resistant','Waterproof']
        df_new['water_resistance_level'] = [values[i % len(values)] for i in range(length)]
    else:
        df_new['water_resistance_level']=['not_water_resistant'] * length

    df_new['is_autographed'] = ['Nein'] * length
    df_new.loc[0, 'is_autographed'] = np.nan
    if 'men' in CiBiao:
        df_new['department'] = ['Herren'] * length
    else:
        df_new['department'] = ['Damen'] * length
    df_new.loc[0, 'department'] = np.nan
    df_new['item_weight_unit_of_measure']=['GR']* length
    weight = df_source.iloc[0:, 7]
    df_new['item_weight']=['150']*length














    df_new['length'] = ['regular']*length
    # key = title[1] + ' ' + word[1]+' ' +  word[2]+' ' +  word[3]


    # print('key:'+str(len(key)))


    for i in range(4):
        # 找到【的索引位置
        index = p5[i].find('【')

        # 如果找到了【，则去除该字符及其之前的部分
        if index != -1:
            p5[i] = p5[i][index:]

    # df_new['care_instructions'] = [p5[0]]*length
    df_new['care_instructions'] = ['Maschinenwäsche'] * length
    df_new['fur_description'] = [p5[2]]*length
    if feed_product == 'pants':
        df_new['fabric_type'] = [p5[3]]*length
    else:
        if 'Linen' in string:
            df_new['fabric_type'] = ['55%linen, 45%cotton'] * length
        else:
            df_new['fabric_type'] = ['71%Polyester,18%Cotton,11%Spandex']*length
    # df_new['import_designation'] = [p5[3]]*length


    product_description = utils.description("html.txt", size_chart, p5)





    # 短埋词区域
    # n = 1
    # for i in range(24):
    #     # word_name = 'word'+str(n)
    #     # df_new[word_name] = [word[n]] * length
    #     df_new['short bury'+str(i)] = np.nan
    #     n=n+1
    # df_new['model1'] = ['Generisch'] * length
    # df_new['model2'] = ['Generisch'] * length
    a=13
    num = length*a


    indices = list(range(len(zi)))

    # 使用random.shuffle打乱索引
    random.shuffle(zi)

    # 使用乱序的索引重新排列Series
    zi = zi.iloc[indices]
    print("共需要",num,"个词")
    print("词表长度为",len(zi))
    if num > len(zi):
        print("词表不够，进行自动扩充")
    while num > len(zi):
        # 使用random.shuffle打乱索引
        random.shuffle(zi)
        indices = list(range(len(zi)))
        # 使用乱序的索引重新排列Series
        zi = zi.iloc[indices]
        zi = pd.concat([zi, zi], ignore_index=True)
        print("词表扩充中...")
        if num < len(zi):
            zi = zi[:num]
            print("词表扩充完毕")
            break
        else:
            continue

    for i in range(a):
        array_name = f"Array{i+1}"  # 创建数组名称
        start_index = i * length  # 计算起始索引
        end_index = (i + 1) * length  # 计算结束索引
        array_data = zi[start_index:end_index].reset_index(drop=True)  # 提取32个元素并重置索引
        df_new[array_name] = array_data  # 将数据填充到新的数据框架的不同列中







    df_new['product_description'] = product_description
    print('大描述长度：'+str(len(product_description)))

    if len(product_description) > 2000:
        print("大描述超过2000字符!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")




    df_new['price'] = [price] * length# 后续要去除第一行
    df_new['quantity'] = ['1000'] * length
    if 'men' in CiBiao:
        df_new['gender'] = ['Männlich'] * length
    else:
        df_new['gender'] = ['Weiblich'] * length
    df_new['age'] = ['Erwachsener'] * length




    if leixing == "pant":
        df_new['size_sys'] = ['DE / NL / SE / PL'] * length
        df_new['Alpha'] = ['Alphanumerisch'] * length
        size_map = [str(string).replace("2X", "XX").replace("3X", "XXX").replace("4X", "XXXX").replace("5X", "XXXXX").replace("6X", "XXXXXX").replace("-Large","L").replace("Large","L").replace("Medium","M").replace("Small","S") for string in size]
        size = [str(string).replace("XX", "2X").replace("XXX", "3X").replace("XXXX", "4X").replace("XXXXX", "5X").replace("XXXXXXX", "6X")for string in size_map]
        df_new['size'] = size
        df_new['body_type'] = ['Regular'] * length
        df_new['height_type'] = ['Regular'] * length
        df_new['size1'] = np.nan
        df_new['size2'] = np.nan
        df_new['size3'] = np.nan
        df_new['size4'] = np.nan
        df_new['size5'] = np.nan
    elif leixing == "coat":
        df_new['size1'] = np.nan
        df_new['size2'] = np.nan
        df_new['size3'] = np.nan
        df_new['size4'] = np.nan
        df_new['size5'] = np.nan
        df_new['size_sys'] = ['DE / NL / SE / PL'] * length
        df_new['Alpha'] = ['Alphanumerisch'] * length
        size_map = [
            str(string).replace("2X", "XX").replace("3X", "XXX").replace("4X", "XXXX").replace("5X", "XXXXX").replace(
                "6X", "XXXXXX").replace("-Large", "L").replace("Large", "L").replace("Medium", "M").replace("Small",
                                                                                                            "S") for
            string in size]
        size = [
            str(string).replace("XX", "2X").replace("XXX", "3X").replace("XXXX", "4X").replace("XXXXX", "5X").replace(
                "XXXXXXX", "6X") for string in size_map]
        df_new['size'] = size
        df_new['body_type'] = ['Regular'] * length
        df_new['height_type'] = ['Regular'] * length




    if zuhe == 1:
        df_new['main_img'] = main_img
    else:
        df_new['main_img'] = main_img
        df_new.loc[:float(count), 'main_img'] = main_img_main

    df_new['img_1'] = img_1
    df_new['img_2'] = img_2
    df_new['img_3'] = img_3
    df_new['img_4'] = img_4
    df_new['img_5'] = img_5
    df_new['img_6'] = final_img
    df_new['img_7'] = img_7
    df_new['final_img'] = img_6
    df_new['swatch_img'] = main_img
    # 清除第1行数据
    df_new.loc[0, ['price', 'quantity', 'age', 'Alpha', 'size_sys', 'body_type', 'height_type', 'main_img', 'img_1', 'img_2', 'img_3', 'img_4', 'img_5', 'img_6', 'img_7', 'final_img','swatch_img']] = np.nan


    df_new['parent_child'] = ['Child'] * length
    df_new.loc[0, 'parent_child'] = 'Parent'
    df_new['parent_sku'] = [sku[0]] * length
    df_new.loc[0, 'parent_sku'] = np.nan
    df_new['relationship'] = ['Variation'] * length
    df_new.loc[0, 'relationship'] = np.nan
    df_new['variation_theme'] = ['color-size'] * length


    df_new['fit_type'] = ['Regular'] * length
    df_new.loc[0, 'fit_type'] = np.nan

    df_new.loc[0, 'rise_style'] = np.nan


    # 创建一个新的数组，将修改后的字符串存储其中

    df_new['size_map'] = size_map
    df_new.loc[0, 'size_map'] = np.nan
    df_new['size_name'] = size_map
    df_new.loc[0, 'size_name'] = np.nan


    df_new['list_price'] = [list_price] * length
    df_new.loc[0, 'list_price'] = np.nan
    df_new['list_price_tax'] = [price] * length
    df_new.loc[0, 'list_price_tax'] = np.nan
    df_new['currency'] = ['EUR'] * length
    df_new.loc[0, 'currency'] = np.nan
    df_new['condition_type'] = ['Neu'] * length
    df_new.loc[0, 'condition_type'] = np.nan


    df_new['shipping'] = [shipping] * length
    df_new.loc[0, 'shipping'] = np.nan

    df_new['package_level'] = ['Einheit'] * length
    df_new['package_contains_quantity'] = ['1'] * length
    # item_length_unit_of_measure	item_length	item_width	item_height	item_width_unit_of_measure	item_height_unit_of_measure	package_height	package_width	package_length	package_length_unit_of_measure	package_weight	package_weight_unit_of_measure	package_height_unit_of_measure	package_width_unit_of_measure
    df_new['item_length_unit_of_measure'] = ['CM'] * length
    df_new['item_length'] = ['25'] * length
    df_new['item_width'] = ['20'] * length
    df_new['item_height'] = ['2.0'] * length
    df_new['item_width_unit_of_measure'] = ['CM'] * length
    df_new['item_height_unit_of_measure'] = ['CM'] * length
    df_new['package_height'] = ['2'] * length
    df_new['package_width'] = ['20'] * length
    df_new['package_length'] = ['25'] * length
    df_new['package_length_unit_of_measure'] = ['CM'] * length
    df_new['package_weight'] = ['150'] * length
    df_new['package_weight_unit_of_measure'] = ['GR'] * length
    df_new['package_height_unit_of_measure'] = ['CM'] * length
    df_new['package_width_unit_of_measure'] = ['CM'] * length
    if leixing == 'jumpsuit' or leixing == 'romper' or leixing == 'overalls':
        df_new['waist_style'] = ['Medium Waist']*length



    # 保存新的数据框架为 Excel 文件
    output_filename = os.path.splitext(os.path.basename(source_file))[0] + '.xlsx'
    output_file = os.path.join(output_filename)
    df_new.to_excel(output_file, index=False)
    file_path = os.path.join(r'F:\pythonproject\pythonProject',output_filename)
    return file_path


