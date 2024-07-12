# 读取 Excel 文件 Data.xlsx 中的工作表 Fruit
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Fruit']

from openpyxl.worksheet.cell_range import CellRange
ws.merge_cells('A1:B2')
ws.merge_cells('D1')
print(ws.merged_cells.ranges)

wb.save('Save.xlsx')

wb1 = load_workbook('Save.xlsx')
ws1 = wb1['Fruit']

print(ws1.merged_cells.ranges)

for c in ws1.merged_cells.ranges:
    print(c)
    print(type(c))
