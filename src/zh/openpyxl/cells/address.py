# 打开 Data.xlsx 中的工作表 Values
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Values']

# 显示 C2 列索引和地址
c2 = worksheet['C2']
print(f'列索引：{c2.column_letter} 地址：{c2.coordinate}')
