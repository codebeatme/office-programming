# 開啟 Data.xlsx 中的工作表 Values
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Values']

b1 = worksheet['B1']
# 取得 B1 左邊的儲存格 A1
print(b1.offset(column=-1))
# 取得 B1 右下角的儲存格 C2
print(b1.offset(1, 1))
