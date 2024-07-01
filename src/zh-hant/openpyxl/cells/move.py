# 開啟 Data.xlsx 中的工作表 Values
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Values']

# B2 將覆蓋 A1
b2 = worksheet['B2']
b2.row = 1
b2.column = 1
# A2 可以順利的移動至 C3
a2 = worksheet['A2']
a2.row = 3
a2.column = 3
workbook.save('Move.xlsx')

# B2 和 A2 在 worksheet 中的位置並沒有改變
print(worksheet['B2'] == b2)
print(worksheet['A2'] == a2)
