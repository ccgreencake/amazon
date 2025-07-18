import Amazon
from datetime import datetime

import utils



def generate_filename(day, prefix,number):
    # 生成目标文件名
    if prefix == "YOULE":
        prefix = "yo"
    elif prefix == "LIANG":
        prefix = "li"
    elif prefix == "SIYAT":
        prefix = "si"

    target_filename = f"{prefix}{day}{number}.xlsx"
    return rf'C:\Users\Administrator\Desktop\{day}\{target_filename}'

def generate_source_filename(day, number):
    # 生成源文件名
    return rf'C:\Users\Administrator\Desktop\product\{day}\product_{number}.xls'

# 根据子标题拼接成ST
fangfa = 0

# 直接根据埋词+属性词拼接成ST
# fangfa = 1

# 拼接子标
zibiao = 0

# 不拼接子标
# zibiao = 1


# 有多个组合图片
zuhe = 1
# 没有组合图片
zuhe = 0

account = "LIANG"
account = "YOULE"
# account = "SIYAT"
# account = "zz"
# account = "lc"



style = 'Classic'


CiBiao = 'men cargo sweatpants'
leixing = 'pants'
TG = -1


price = 8.99
shipping = 4.99
count = 0
all_path = []
item_type = 'athletic-pants'
manufacturer = 'Bakgeerle'
for i in range(25):
    no = 2
    color_number = i+no
    #正常上架
    # color_number = 10
    num = color_number
    if num < 10:
        number = "0"+str(num)
    else:
        number = str(num)
    if TG == -1:
        TG_VP = 'P'
        if count % 5 == 0:
            price = 1 + price
            shipping = shipping - 1
        count+=1
    elif i == 0 and TG == 1:
        TG_VP = 'T'
        price = 32.49
        shipping = 0
    else:
        TG_VP = 'C'
        price = i + 19.99
        shipping = 6.99 - i


    # TG_VP = 'T'
    # price = 37.49
    # shipping = 0

    # 获取当前日期并格式化为YYYYMMDD的格式
    current_date_str = datetime.now().strftime('%Y%m%d')
    # 截取字符串，去掉前两位
    day = current_date_str[2:]
    # day = "240820"
    # 使用函数生成文件路径
    mubiao_file = generate_filename(day, account, number)
    source_file = generate_source_filename(day,number)
    file_path = Amazon.amazon(source_file,CiBiao,account,TG_VP,price,shipping,leixing,style,fangfa,color_number,zuhe,zibiao,item_type,manufacturer)
    if TG_VP == 'P':
        all_path.append(file_path)
    else:
        utils.copy_file(mubiao_file, file_path, TG_VP, color_number,account)

if TG == -1:
    current_date_str = datetime.now().strftime('%Y%m%d')
    day = current_date_str[2:]
    mubiao_file = generate_filename(day, account, '02')
    utils.copy_file(mubiao_file, all_path, 'P',2,account)
