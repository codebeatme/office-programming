# 读取 Excel 文件 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Trees']

from openpyxl.worksheet.cell_range import CellRange
# 创建区域 A1:B2
range = CellRange('A1:B2')

# 借助 CellRange 的 row 属性遍历单元格
for row in range.rows:
    # x 和 y 分别表示单元格位于哪一行和哪一列
    for x, y in row:
        c = ws.cell(x, y)
        print(f'{c.coordinate}={c.value}')

# 借助 CellRange 的 cells 属性遍历单元格
for x, y in range.cells:
    c = ws.cell(x, y)
    print(f'({x}, {y})={c.value}')