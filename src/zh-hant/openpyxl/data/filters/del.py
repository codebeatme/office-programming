# 開啟 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Teachers']

# 選出與第二欄有關的所有 FilterColumn 物件
fcs = worksheet.auto_filter.filterColumn
del_fcs = [x for x in fcs if x.colId == 1]

# 移除與第二欄有關的 FilterColumn 物件
for fc in del_fcs:
    fcs.remove(fc)

workbook.save('Del.xlsx')
