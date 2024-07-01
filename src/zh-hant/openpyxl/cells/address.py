# 開啟 Data.xlsx 中的工作表 Values
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Values']

# 顯示 C2 欄索引和位址
c2 = worksheet['C2']
print(f'欄索引：{c2.column_letter} 位址：{c2.coordinate}')
