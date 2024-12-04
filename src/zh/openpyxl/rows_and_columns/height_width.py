# 创建只写工作簿，并添加工作表 HW
from openpyxl import Workbook
wb = Workbook(True)
ws = wb.create_sheet('HW')

# 设置第一行的高度
ws.row_dimensions[1].height = 30
# 设置第一列的宽度
ws.column_dimensions['A'].width = 30

wb.save('HW.xlsx')

# 是否拥有自定义高度和宽度？
print(ws.row_dimensions[1].customHeight)
print(ws.column_dimensions['A'].customWidth)