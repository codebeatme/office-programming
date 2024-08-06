# 開啟 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Trees']

# 設定對 A，B 兩欄進行自動篩選
worksheet.auto_filter.ref = 'A1:B1'

workbook.save('Ref.xlsx')