import Amazon
from datetime import datetime
def generate_filename(day, prefix,number):
    # 生成目标文件名
    if prefix == "YOULE":
        prefix = "yo"
    elif prefix == "LIANG":
        prefix = "li"
    target_filename = f"{prefix}{day}{number}.xlsx"
    return rf'C:\Users\123\Desktop\{day}\{target_filename}'

def generate_source_filename(day, number):
    # 生成源文件名
    return rf'C:\Users\123\Desktop\product\{day}\product_{number}.xls'

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
# account = "lc"







# CiBiao = 'sweater'
# style = 'Casual'
# leixing = 'sweater'

# CiBiao = 'corduroy jacket'
# style = 'Casual'
# leixing = 'coat'
#
# CiBiao = 'xmas cardigan'
# style = 'Casual'
# leixing = 'sweater'
# CiBiao = 'wide_leg_yoga'
# leixing = 'pants'

style = 'Casual'

# CiBiao = 'men tactical hoodie'
# style = 'Casual'
# leixing = 'sweatshirt'

# CiBiao = 'plaid skirt'
# style = 'Casual'
# leixing = 'shirt'

CiBiao = 'weiku'
style = 'Casual'
leixing = 'pants'


CiBiao = 'men tactical hoodie'
leixing = 'sweatshirt'




TG_VP = 'C'

# price = '14.99'
# shipping = '1.99'

for i in range(2):
    color_number = i+1
    #正常上架
    # color_number = 10
    number = "0"+str(color_number+1)
    price = i+8.99
    shipping = 9.99-i

    # 获取当前日期并格式化为YYYYMMDD的格式
    current_date_str = datetime.now().strftime('%Y%m%d')
    # 截取字符串，去掉前两位
    day = current_date_str[2:]
    # day = "240820"
    # 使用函数生成文件路径
    mubiao_file = generate_filename(day, account, number)
    source_file = generate_source_filename(day,number)
    Amazon.amazon(mubiao_file,source_file,CiBiao,account,TG_VP,price,shipping,leixing,style,fangfa,color_number,zuhe,zibiao)
