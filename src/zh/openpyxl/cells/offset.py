# 打开 Data.xlsx 中的工作表 Values
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Values']

b1 = worksheet['B1']
# 获取 B1 左边的单元格 A1
print(b1.offset(column=-1))
# 获取 B1 右下角的单元格 C2
print(b1.offset(1, 1))
