# 開啟 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Trees']

from openpyxl.worksheet.filters import FilterColumn, Filters
af = ws.auto_filter

# 為第一欄新增自動篩選
f = Filters(filter=('蘋果樹', '梨樹'))
# 將覆蓋之前在 Filters 中設定的 filter
fc = FilterColumn(0, filters=f, blank=True, vals=('大樹', '柳樹'))
af.filterColumn.append(fc)

# 選出第四欄中值為 3，3.5 的儲存格或空白儲存格
af.add_filter_column(3, [3, 3.5], True)

# 將第一和第四欄包含在進行自動篩選的範圍中
af.ref = 'A1:D1'
# 在儲存的 Excel 檔案中，可能需要重新應用自動篩選才能看到效果
wb.save('Add.xlsx')