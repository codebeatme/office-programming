# 打开 Food.xlsx 中的工作表 Tea
import openpyxl
workbook = openpyxl.load_workbook('Food.xlsx')
worksheet = workbook['Tea']

# 当工作表被保护时，将隐藏第一行单元格的公式，并锁定第一行单元格
col = worksheet.column_dimensions['B']
col.reindex()
print(col.range)