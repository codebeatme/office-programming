# 開啟 Data.xlsx 中的工作表 Teachers
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Teachers']

# 選出與第二欄有關的所有 FilterColumn 物件
fcs = worksheet.auto_filter.filterColumn
del_fcs = [x for x in fcs if x.colId == 1]

# 移除與第二欄有關的 FilterColumn 物件
for fc in del_fcs:
    fcs.remove(fc)

# 修改第一欄的 FilterColumn 物件
for fc in fcs:
    if fc.colId == 0:
        fc.filters.filter = ('大剛', '小剛')

# 在儲存的 Excel 檔案中，可能需要重新應用自動篩選才能看到效果
workbook.save('Change.xlsx')
