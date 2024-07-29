# 讀取 Excel 檔案 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Trees']

from openpyxl.worksheet.cell_range import CellRange
# 建立範圍 A1:B2
range = CellRange('A1:B2')

# 借助 CellRange 的 row 屬性周遊儲存格
for row in range.rows:
    # x 和 y 分別表示儲存格位於哪一列和哪一欄
    for x, y in row:
        c = ws.cell(x, y)
        print(f'{c.coordinate}={c.value}')

# 借助 CellRange 的 cells 屬性周遊儲存格
for x, y in range.cells:
    c = ws.cell(x, y)
    print(f'({x}, {y})={c.value}')