# 開啟 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Trees']

af = worksheet.auto_filter

# 選出第二欄中值為 20 的儲存格或空白儲存格
af.add_filter_column(1, [20], True)
# 選出第二欄中值為 15 的儲存格
af.add_filter_column(1, [15], False)

workbook.save('AddMulti.xlsx')
