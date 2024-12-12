import openpyxl
import random
import subprocess
import re

excel_file = r'C:\Users\123\Desktop\new 词表.xlsx'

# 读取Excel文件
wb = openpyxl.load_workbook(excel_file)
shuxing = [
"baggy",
"2024",
"casual",
"plus size",
"comfy ",
"soft ",
"comfortable ",
"elegant  ",
"cozy ",
"trendy ",
"loose fit",
"winter",
"fall"
]
# 选择名为 "barrel_jeans" 的工作表

# 要处理的Excel文件路径
def zibiao(num,cibiao,title):
    ws = wb[cibiao]
    changdu = 200
    texts = [cell.value.strip() for cell in ws['C'] if cell.value and isinstance(cell.value, str)]
    texts.pop(0)
    if len(texts) < 1:
        print("工作表中没有足够的文本数据。")
    else:
        # 生成num个长度接近250字符的随机字符串
        tiaoguo = 0
        random_strings = []
        for _ in range(num):
            random_string = title[_]+" "
            total_length = len(random_string)
            flag = True
            if tiaoguo == 1:
                while total_length < changdu and flag:  # 当总长度小于200且还有文本片段时继续
                    for text in random.sample(texts, len(texts)):  # 随机打乱texts顺序
                        if total_length + len(text) + 1 <= changdu:  # +1是为了加上空格
                            random_string += text + " "  # 添加文本片段和空格
                            total_length += len(text) + 1
                        else:
                            flag = False
                            break  # 退出内层循环，保持当前random_string
                random_string = random_string.strip()
                # 第一次处理重复单词
                random_string = re.sub(r'\b(\w+)\s+\1\b', r'\1', random_string)
                flag = True
                while total_length < changdu and flag:  # 当总长度小于200且还有文本片段时继续
                    for text in random.sample(texts, len(texts)):  # 随机打乱texts顺序
                        if total_length + len(text) + 1 <= changdu:  # +1是为了加上空格
                            random_string += text + " "  # 添加文本片段和空格
                            total_length += len(text) + 1
                        else:
                            flag = False
                            break  # 退出内层循环，保持当前random_string
                random_string = re.sub(r'\b(\w+)\s+\1\b', r'\1', random_string)
            else:
                tiaoguo = 1
            # 去除最后一个多余的空格
            random_strings.append(random_string)
            random_string = random_string.strip()
        return random_strings

def pinjie1(num,cibiao,title,fangfa):
    # excel_file = r'C:\Users\123\Desktop\new 词表.xlsx'
    #
    # # 读取Excel文件
    # wb = openpyxl.load_workbook(excel_file)

    ws = wb[cibiao]
    changdu = 250
    # changdu=500
    # 读取第一列的所有文本
    texts = [cell.value.strip() for cell in ws['A'] if cell.value and isinstance(cell.value, str)]
    texts.pop(0)
    # 检查texts列表是否有足够的文本片段
    if len(texts) < 1:
        print("工作表中没有足够的文本数据。")
    else:
        # 生成num个长度接近250字符的随机字符串
        random_strings = []
        for _ in range(num):
            if fangfa == 0:
                random_string = title[_]+" "
                total_length = len(random_string)
            else :
                random_string = ""
                total_length = 0
            flag = True
            while total_length < changdu and flag :  # 当总长度小于250且还有文本片段时继续
                shuxing = [
                    "baggy",
                    "casual",
                    "plus size",
                    "comfy",
                    "soft",
                    "comfortable",
                    "elegant",
                    "cozy",
                    "trendy",
                    "loose fit",
                    "winter",
                    "fall",
                    "fleece",
                    "oversized",
                    "petite",
                    "y2k"
                ]
                random.shuffle(shuxing)
                shuxing = shuxing+shuxing[:]+shuxing[:]+shuxing[:]
                n=0
                for text in random.sample(texts, len(texts)):  # 随机打乱texts顺序
                    if total_length + len(text) + 3 + len(shuxing[n]) + len(shuxing[n+1]) <= changdu:  # +3是为了加上空格
                        random_string += shuxing[n] + " " + shuxing[n+1] + " " + text + " "  # 添加文本片段和空格
                        total_length += len(text) + 3 + len(shuxing[n]) + len(shuxing[n+1])
                        n+=2
                    else:
                        flag= False
                        break  # 退出内层循环，保持当前random_string

            # 去除最后一个多余的空格
            random_string = random_string.strip()
            random_strings.append(random_string)
    return random_strings
        # # 创建一个新的工作簿来保存结果
        # new_wb = openpyxl.Workbook()
        # new_ws = new_wb.active
        #
        # # 将生成的字符串及其长度写入新的工作表
        # for i, random_string in enumerate(random_strings, 1):
        #     new_ws.cell(row=i, column=1, value=random_string)  # 写入字符串
        #     new_ws.cell(row=i, column=2, value=len(random_string))  # 写入字符串长度
        #
        # # 定义新Excel文件的保存路径
        # new_excel_file = r'C:\Users\123\Desktop\random_strings.xlsx'
        # # 保存新的Excel文件
        # new_wb.save(new_excel_file)
        #
        # print(f"文件已保存: {new_excel_file}")
        #
        # # 使用WPS打开Excel文件（如果需要）
        # wps_path = r'C:\Users\123\AppData\Local\Kingsoft\WPS Office\ksolaunch.exe'
        # subprocess.call([wps_path, new_excel_file])
        # print(f"文件 '{new_excel_file}' 已使用WPS打开。")
def pinjie2(num, cibiao, flag):
    ws = wb[cibiao]
    changdu = 500
    texts = [cell.value.strip() for cell in ws['C'] if cell.value and isinstance(cell.value, str)]
    texts.pop(0)
    if len(texts) < 1:
        print("工作表中没有足够的文本数据。")
    else:
        random_strings = []
        for _ in range(num):
            if flag == 'li':
                random_string = "Bakgeerle "
                total_length = 10
            elif flag == 'lc':
                random_string = "lcyhony "
                total_length = 8
            else:
                random_string = ""
                total_length = 0
            flag = True
            while total_length < changdu and flag:
                for text in random.sample(texts, len(texts)):
                    if total_length + len(text) + 1 <= changdu:
                        random_string += text + " "
                        total_length += len(text) + 1
                    else:
                        flag = False
                        break
            random_string = random_string.strip()
            # 第一次处理重复单词
            random_string = re.sub(r'\b(\w+)\s+\1\b', r'\1', random_string)
            flag = True
            # 检查长度是否小于500，如果是，则继续拼接
            if len(random_string) < changdu:
                while total_length < changdu and flag:
                    for text in random.sample(texts, len(texts)):
                        if len(random_string) + len(text) + 1 <= changdu:
                            random_string += " " + text
                            total_length += len(text) + 1
                        else:
                            flag = False
                            break
            random_string = random_string.strip()
            # 第二次处理重复单词
            random_string = re.sub(r'\b(\w+)\s+\1\b', r'\1', random_string)
            random_strings.append(random_string)
    return random_strings

def pinjiedaochu(num,cibiao,flag):
    ws = wb[cibiao]
    # changdu=250
    changdu = 500
    # 读取第一列的所有文本
    texts = [cell.value.strip() for cell in ws['A'] if cell.value and isinstance(cell.value, str)]
    texts.pop(0)
    # 检查texts列表是否有足够的文本片段
    if len(texts) < 1:
        print("工作表中没有足够的文本数据。")
    else:
        # 生成num个长度接近250字符的随机字符串
        random_strings = []
        for _ in range(num):
            if flag == 'li':
                random_string = "Bakgeerle "
                total_length = 10
            elif flag == 'lc':
                random_string = "lcyhony "
                total_length = 8
            else:
                random_string = ""
                total_length = 0
            flag = True
            while total_length < changdu and flag:  # 当总长度小于250且还有文本片段时继续
                for text in random.sample(texts, len(texts)):  # 随机打乱texts顺序
                    if total_length + len(text) + 1 <= changdu:  # +1是为了加上空格
                        random_string += text + " "  # 添加文本片段和空格
                        total_length += len(text) + 1
                    else:
                        flag = False
                        break  # 退出内层循环，保持当前random_string

            # 去除最后一个多余的空格
            random_string = random_string.strip()
            random_strings.append(random_string)
            # 创建一个新的工作簿来保存结果
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active

    # 将生成的字符串及其长度写入新的工作表
    for i, random_string in enumerate(random_strings, 1):
        new_ws.cell(row=i, column=1, value=random_string)  # 写入字符串
        new_ws.cell(row=i, column=2, value=len(random_string))  # 写入字符串长度

    # 定义新Excel文件的保存路径
    new_excel_file = r'C:\Users\123\Desktop\random_strings.xlsx'
    # 保存新的Excel文件
    new_wb.save(new_excel_file)

    print(f"文件已保存: {new_excel_file}")

    # 使用WPS打开Excel文件（如果需要）
    wps_path = r'C:\Users\123\AppData\Local\Kingsoft\WPS Office\ksolaunch.exe'
    subprocess.call([wps_path, new_excel_file])
    print(f"文件 '{new_excel_file}' 已使用WPS打开。")
# pinjiedaochu(10,'weiku','li')