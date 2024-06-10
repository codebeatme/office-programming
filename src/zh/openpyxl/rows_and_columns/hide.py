# 读取 Excel 文件 Food.xlsx 中的工作表 Fruit
from openpyxl import load_workbook
workbook = load_workbook('Food.xlsx')
worksheet = workbook['Fruit']

# 隐藏第一行或第一列
worksheet.row_dimensions[1].hidden = True
worksheet.column_dimensions['A'].hidden = True

workbook.save('Hidden.xlsx')
