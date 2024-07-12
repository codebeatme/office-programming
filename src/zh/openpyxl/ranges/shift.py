from openpyxl.worksheet.cell_range import CellRange

# 创建区域 B2:D4
range = CellRange('B2:D4')
# 向右下方移动区域
range.shift(1, 1)
print(range.coord)

# ERROR 移动后区域将超出工作表的范围
range.shift(row_shift=-3)