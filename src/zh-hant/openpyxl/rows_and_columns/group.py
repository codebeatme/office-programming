# 讀取 Excel 檔案 Food.xlsx 中的工作表 Sweets
from openpyxl import load_workbook
workbook = load_workbook('Food.xlsx')
worksheet = workbook['Sweets']

# 組合工作表的區域 B2:D4
worksheet.row_dimensions.group(2, 4, hidden=True)
worksheet.column_dimensions['B'].outlineLevel = 1
worksheet.column_dimensions['C'].outlineLevel = 2
worksheet.column_dimensions['D'].outline_level = 1

workbook.save('Group.xlsx')
