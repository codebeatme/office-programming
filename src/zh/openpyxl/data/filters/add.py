# 打开 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Trees']

from openpyxl.worksheet.filters import FilterColumn, Filters
af = ws.auto_filter

# 为第一列添加自动筛选
f = Filters(filter=('苹果树', '梨树'))
# 将覆盖之前在 Filters 中设置的 filter
fc = FilterColumn(0, filters=f, blank=True, vals=('大树', '柳树'))
af.filterColumn.append(fc)

# 选出第四列中值为 3，3.5 的单元格或空单元格
af.add_filter_column(3, [3, 3.5], True)

# 将第一和第四列包含在进行自动筛选的区域中
af.ref = 'A1:D1'
# 在保存的 Excel 文件中，可能需要重新应用自动筛选才能看到效果
wb.save('Add.xlsx')