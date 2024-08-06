# 打开 Data.xlsx 中的工作表 Games
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Games']

from openpyxl.worksheet.filters import FilterColumn, CustomFilter, CustomFilters

# 选出数字内容大于 9 并且 小于等于 11 的 Excel 单元格
cfs = CustomFilters(True, [
    CustomFilter('greaterThan', '9'),
    CustomFilter('lessThanOrEqual', '11')
    ])

# 为工作表的第二列设置自定义筛选
fc = FilterColumn(1, customFilters=cfs)
ws.auto_filter.filterColumn.append(fc)

# 选出第一列中文字内容不等于 疯狂自行车 的 Excel 单元格
cfs = CustomFilters(customFilter=[CustomFilter('notEqual', '疯狂自行车')])
ws.auto_filter.filterColumn.append(FilterColumn(0, customFilters=cfs))

# 在保存的 Excel 文件中，可能需要重新应用自动筛选才能看到效果
wb.save('Custom.xlsx')