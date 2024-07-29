# 讀取 Excel 檔案 Data.xlsx 中的工作表 Fruit
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Fruit']

from openpyxl.worksheet.cell_range import CellRange
# 建立範圍 Trees!A1:B2，其中 Trees 不會發揮作用
range = CellRange(min_col=1, min_row=1, max_col=2, max_row=2, title='Trees')
# 移動工作表 Fruit 的範圍 A1:B2，而不是工作表 Trees
ws.move_range(range, 1, 1)

wb.save('Move.xlsx')