# 建立唯寫活頁簿，並新增工作表 HW
from openpyxl import Workbook
wb = Workbook(True)
ws = wb.create_sheet('HW')

# 設定第一列的高度
ws.row_dimensions[1].height = 30
# 設定第一欄的寬度
ws.column_dimensions['A'].width = 30

wb.save('HW.xlsx')

# 是否擁有自訂高度和寬度？
print(ws.row_dimensions[1].customHeight)
print(ws.column_dimensions['A'].customWidth)