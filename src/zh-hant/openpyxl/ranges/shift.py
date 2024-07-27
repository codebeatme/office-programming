from openpyxl.worksheet.cell_range import CellRange

# 建立區域 B2:D4
range = CellRange('B2:D4')
# 向右下方移動區域
range.shift(1, 1)
print(range.coord)

# ERROR 移動後區域將超出工作表的範圍
range.shift(row_shift=-3)