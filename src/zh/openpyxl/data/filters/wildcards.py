# 打开 Data.xlsx 中的工作表 Games
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Games']

from openpyxl.worksheet.filters import FilterColumn, CustomFilters, CustomFilter, StringFilter

# 选出第一列中文字内容包含 自?车 的单元格
cfs = CustomFilters(customFilter=[CustomFilter('equal', '*自?车*')])
ws.auto_filter.filterColumn.append(FilterColumn(0, customFilters=cfs))

# 选出第二列中包含个位数的单元格
cfs = CustomFilters(customFilter=[StringFilter('wildcard', '?')])
ws.auto_filter.filterColumn.append(FilterColumn(1, customFilters=cfs))

# 在保存的 Excel 文件中，可能需要重新应用自动筛选才能看到效果
wb.save('Wildcards.xlsx')