# 打开 Data.xlsx 中的工作表 Types
import openpyxl
workbook = openpyxl.load_workbook('Data.xlsx')
worksheet = workbook['Types']

from openpyxl.styles.protection import Protection

g1 = worksheet['G1']
# 显示单元格 G1 的保护信息
print(g1.protection)
# 当工作表被保护时，将隐藏 G1 单元格的公式
g1.protection = Protection(False, True)

workbook.save('P.xlsx')