# 開啟 Data.xlsx 中的工作表 Students
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Trees']

# print(worksheet.auto_filter.filterColumn)

for c in worksheet.auto_filter.filterColumn:
    print(c)

# worksheet.auto_filter.add_filter_column()