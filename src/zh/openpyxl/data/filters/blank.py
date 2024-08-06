# 打开 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Trees']

from openpyxl.worksheet.filters import FilterColumn, CustomFilters, BlankFilter

# 选出第一列中文字内容不包含 不 的 Excel 单元格
cfs = CustomFilters(customFilter=[BlankFilter()])
ws.auto_filter.filterColumn.append(FilterColumn(0, customFilters=cfs))

# 在保存的 Excel 文件中，可能需要重新应用自动筛选才能看到效果
wb.save('Blank.xlsx')