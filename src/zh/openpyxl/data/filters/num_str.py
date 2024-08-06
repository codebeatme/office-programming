# 打开 Data.xlsx 中的工作表 Games
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Games']

from openpyxl.worksheet.filters import FilterColumn, CustomFilters, NumberFilter, StringFilter

# 选出数字内容大于等于 10 或 小于等于 11 的 Excel 单元格
cfs = CustomFilters(True, [
    NumberFilter('greaterThanOrEqual', 10),
    NumberFilter('lessThanOrEqual', 11)
    ])

# 为工作表的第二列设置自定义筛选
fc = FilterColumn(1, customFilters=cfs)
ws.auto_filter.filterColumn.append(fc)

# 选出第一列中文字内容不包含 不 的 Excel 单元格
cfs = CustomFilters(customFilter=[StringFilter('contains', '不', True)])
ws.auto_filter.filterColumn.append(FilterColumn(0, customFilters=cfs))

# 在保存的 Excel 文件中，可能需要重新应用自动筛选才能看到效果
wb.save('NumStr.xlsx')