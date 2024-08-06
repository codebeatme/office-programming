# 開啟 Data.xlsx 中的工作表 Games
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Games']

from openpyxl.worksheet.filters import FilterColumn, CustomFilters, NumberFilter, StringFilter

# 選出數值內容大於等於 10 或 小於等於 11 的 Excel 儲存格
cfs = CustomFilters(True, [
    NumberFilter('greaterThanOrEqual', 10),
    NumberFilter('lessThanOrEqual', 11)
    ])

# 為工作表的第二欄設定自訂篩選
fc = FilterColumn(1, customFilters=cfs)
ws.auto_filter.filterColumn.append(fc)

# 選出第一欄中文字內容不包含 不 的 Excel 儲存格
cfs = CustomFilters(customFilter=[StringFilter('contains', '不', True)])
ws.auto_filter.filterColumn.append(FilterColumn(0, customFilters=cfs))

# 在儲存的 Excel 檔案中，可能需要重新應用自動篩選才能看到效果
wb.save('NumStr.xlsx')