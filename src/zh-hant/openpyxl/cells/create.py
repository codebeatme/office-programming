# 開啟 Data.xlsx 中的工作表 Values
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Values']

from datetime import date
from openpyxl import Workbook
from openpyxl.cell.cell import Cell, WriteOnlyCell

# 建立 Cell 並新增至工作表物件
ws.append((Cell(ws, 1, 1, '儲存格 A1'), 1.23, '2024-11-11'))
ws.append((Cell(ws, value='儲存格 B1'), 4.56, date(2024, 1, 1)))
Cell(ws, 5, 5, value='我是不會被新增的')
# 儲存至 Excel 檔案 Add.xlsx
wb.save('Add.xlsx')

w_ws = Workbook(True).create_sheet()
# 建立 Cell 並新增至唯寫工作表物件
w_ws.append([Cell(w_ws, 2, 2, '2 2'), Cell(w_ws, 1, 1, '1 1')])
w_ws.append([WriteOnlyCell(w_ws, '唯寫')])
# 儲存至 Excel 檔案 New.xlsx
w_ws.parent.save('New.xlsx')