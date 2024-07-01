# 打开 Data.xlsx 中的工作表 Values
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Values']

from datetime import date
from openpyxl import Workbook
from openpyxl.cell.cell import Cell, WriteOnlyCell

# 创建 Cell 并添加至工作表对象
ws.append((Cell(ws, 1, 1, '单元格 A1'), 1.23, '2024-11-11'))
ws.append((Cell(ws, value='单元格 B1'), 4.56, date(2024, 1, 1)))
Cell(ws, 5, 5, value='我是不会被添加的')
# 保存至 Excel 文件 Add.xlsx
wb.save('Add.xlsx')

w_ws = Workbook(True).create_sheet()
# 创建 Cell 并添加至只写工作表对象
w_ws.append([Cell(w_ws, 2, 2, '2 2'), Cell(w_ws, 1, 1, '1 1')])
w_ws.append([WriteOnlyCell(w_ws, '只写')])
# 保存至 Excel 文件 New.xlsx
w_ws.parent.save('New.xlsx')