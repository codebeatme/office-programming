# 读取 Excel 文件 Food.xlsx 中的工作表 Sweets
from openpyxl import load_workbook
workbook = load_workbook('Food.xlsx')
worksheet = workbook['Sweets']

# 按照行的方式遍历单元格区域 B1:C3
for row in worksheet.iter_rows(max_row=3, min_col=2, max_col=3):
    print(row)

# 按照列的方式遍历区域 A1:B2 内的单元格的值
for column_values in worksheet.iter_cols(max_col=2, max_row=2, values_only=True):
    print(column_values)
