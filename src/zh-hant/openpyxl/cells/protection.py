# 開啟 Data.xlsx 中的工作表 Types
import openpyxl
workbook = openpyxl.load_workbook('Data.xlsx')
worksheet = workbook['Types']

from openpyxl.styles.protection import Protection

g1 = worksheet['G1']
# 顯示儲存格 G1 的保護資訊
print(g1.protection)
# 當工作表被保護時，將隱藏 G1 儲存格的公式
g1.protection = Protection(False, True)

workbook.save('P.xlsx')