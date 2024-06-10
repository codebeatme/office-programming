# 读取 Excel 文件 Food.xlsx 中的工作表 Fruit
from openpyxl import load_workbook
workbook = load_workbook('Food.xlsx')
worksheet = workbook['Fruit']

# 工作表最大行是 2，最大列是 2
for row in worksheet.rows:
    print(row)

for column in worksheet.columns:
    print(column)
