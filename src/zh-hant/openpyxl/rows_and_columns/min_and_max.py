# 讀取 Excel 檔案 Food.xlsx 中的工作表 Cakes
from openpyxl import open
workbook = open('Food.xlsx')
worksheet = workbook['Cakes']

# 工作表的已用儲存格的最小區域為 B2:C3
print(f'最小列 {worksheet.min_row}，最小欄 {worksheet.min_column}，最大列 {worksheet.max_row}，最大欄 {worksheet.max_row}')

# 在存取儲存格 A1 和 D4 之後，最小區域發生改變
worksheet['A1']
worksheet['D4']
print(f'最小列 {worksheet.min_row}，最小欄 {worksheet.min_column}，最大列 {worksheet.max_row}，最大欄 {worksheet.max_row}')
