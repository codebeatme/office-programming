from openpyxl.worksheet.cell_range import CellRange

# 建立範圍 B2:D4
range = CellRange('B2:D4')
# 向右下方移動範圍
range.shift(1, 1)
print(range.coord)

# ERROR 移動後範圍將超出工作表的範圍
range.shift(row_shift=-3)