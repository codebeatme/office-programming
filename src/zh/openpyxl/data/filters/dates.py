# 打开 Data.xlsx 中的工作表 Games
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Dates']

from openpyxl.worksheet.filters import FilterColumn, Filters, DateGroupItem

fc = FilterColumn(0, filters=Filters(dateGroupItem=(
    DateGroupItem(month=1, dateTimeGrouping='month'),
    DateGroupItem(year=2024, dateTimeGrouping='year'),
    )))
ws.auto_filter.filterColumn.append(fc)

# 在保存的 Excel 文件中，可能需要重新应用自动筛选才能看到效果
wb.save('Date.xlsx')