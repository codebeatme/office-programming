# 打开 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Trees']

af = worksheet.auto_filter

# 选出第二列中值为 20 的单元格或空单元格
af.add_filter_column(1, [20], True)
# 选出第二列中值为 15 的单元格
af.add_filter_column(1, [15], False)

workbook.save('AddMulti.xlsx')
