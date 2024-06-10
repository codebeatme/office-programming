# 讀取 Excel 檔案 Food.xlsx 中的工作表 Fruit
from openpyxl import load_workbook
workbook = load_workbook('Food.xlsx')
worksheet = workbook['Fruit']

# 隱藏第一列或第一欄
worksheet.row_dimensions[1].hidden = True
worksheet.column_dimensions['A'].hidden = True

workbook.save('Hidden.xlsx')
