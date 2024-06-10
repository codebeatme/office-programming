# 读取 Excel 文件 Food.xlsx 中的工作表 Fruit
from openpyxl import load_workbook
workbook = load_workbook('OL.xlsx')
worksheet = workbook['Fruit']

# from openpyxl import Workbook
# workbook = Workbook(True)
# worksheet = workbook.create_sheet()

# 隐藏第一行或第一列
print(worksheet.column_dimensions['B'].index)
print(worksheet.row_dimensions[1].index)
print(workbook['Fruit'].row_dimensions[1].index)
# workbook.save('OL.xlsx')
