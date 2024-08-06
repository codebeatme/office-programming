# 開啟 Data.xlsx 中的工作表 Students
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Trees']

print(worksheet.auto_filter.filterColumn)

# for c in worksheet.auto_filter.filterColumn:
    # print(c)

# print(worksheet.auto_filter.filterColumn[0].colorFilter)

# worksheet.auto_filter.add_filter_column()

from openpyxl.worksheet.cell_range import CellRange

worksheet.auto_filter.ref = 'A1:A3'

workbook.save('F.xlsx')