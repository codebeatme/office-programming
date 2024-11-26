# 開啟 Data.xlsx 中的工作表 Games
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Games']

from openpyxl.worksheet.filters import FilterColumn, CustomFilters, CustomFilter, StringFilter

# 選出第一欄中文字內容包含 自?車 的儲存格
cfs = CustomFilters(customFilter=[CustomFilter('equal', '*自?車*')])
ws.auto_filter.filterColumn.append(FilterColumn(0, customFilters=cfs))

# 選出第二欄中包含個位數的儲存格
cfs = CustomFilters(customFilter=[StringFilter('wildcard', '?')])
ws.auto_filter.filterColumn.append(FilterColumn(1, customFilters=cfs))

# 在儲存的 Excel 檔案中，可能需要重新應用自動篩選才能看到效果
wb.save('Wildcards.xlsx')