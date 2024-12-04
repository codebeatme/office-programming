# 打开 Food.xlsx 中的工作表 Sandwich
import openpyxl
workbook = openpyxl.load_workbook('Food.xlsx')
worksheet = workbook['Sandwich']

from openpyxl.styles.protection import Protection

# 当工作表被保护时，将隐藏第一行单元格的公式，并锁定第一行单元格
worksheet.row_dimensions[1].protection = Protection(True, True)

workbook.save('PRow.xlsx')