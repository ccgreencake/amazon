from 复制 import copyfile2

n = 19
account = "yo"
# account = "li"
# account = "zz"
if account == "yo":
    path = 'yo总模板.xlsx'
elif account == "li":
    path = 'li总模板.xlsx'
elif account == "zz":
    path = 'zz模板.xlsx'

# copyfile1(n,account,path)
path = r"F:\product\260302\product_04.xls"
copyfile2(path,n)