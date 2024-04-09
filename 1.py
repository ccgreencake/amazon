import pandas as pd



source_file2 = r'C:\Users\123\Desktop\new 词表.xlsx'  # 替换为实际的源文件路径
df_source2 = pd.read_excel(source_file2, sheet_name='cargo')
point5 = df_source2.iloc[0:, 1]
print(point5)