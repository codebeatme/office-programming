# 打开 Data.xlsx 中的工作表 Teachers
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Teachers']

# 选出与第二列有关的所有 FilterColumn 对象
fcs = worksheet.auto_filter.filterColumn
del_fcs = [x for x in fcs if x.colId == 1]

# 移除与第二列有关的 FilterColumn 对象
for fc in del_fcs:
    fcs.remove(fc)

# 修改第一列的 FilterColumn 对象
for fc in fcs:
    if fc.colId == 0:
        fc.filters.filter = ('大刚', '小刚')

# 在保存的 Excel 文件中，可能需要重新应用自动筛选才能看到效果
workbook.save('Change.xlsx')
