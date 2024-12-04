# 開啟 Food.xlsx 中的工作表 Sandwich
import openpyxl
workbook = openpyxl.load_workbook('Food.xlsx')
worksheet = workbook['Sandwich']

from openpyxl.styles.protection import Protection

# 當工作表被保護時，將隱藏第一列儲存格的公式，並鎖定第一列儲存格
worksheet.row_dimensions[1].protection = Protection(True, True)

workbook.save('PRow.xlsx')