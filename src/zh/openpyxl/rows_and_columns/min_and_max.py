# 读取 Excel 文件 Food.xlsx 中的工作表 Cakes
from openpyxl import open
workbook = open('Food.xlsx')
worksheet = workbook['Cakes']

# 工作表的已用单元格的最小区域为 B2:C3
print(f'最小行 {worksheet.min_row}，最小列 {worksheet.min_column}，最大行 {worksheet.max_row}，最大列 {worksheet.max_row}')

# 在访问单元格 A1 和 D4 之后，最小区域发生改变
worksheet['A1']
worksheet['D4']
print(f'最小行 {worksheet.min_row}，最小列 {worksheet.min_column}，最大行 {worksheet.max_row}，最大列 {worksheet.max_row}')
