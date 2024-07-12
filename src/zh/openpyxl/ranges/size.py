from openpyxl.worksheet.cell_range import CellRange

# 创建区域 B2:D4
range = CellRange('B2:D4')
# 显示区域大小
print(range.size)

# 最右边列向内移动 1 列，最下边行向外移动 1 行，最左边列向外移动 1 列，最上边行向内移动 1 行
range.expand(-1, 1, 1, -1)
print(range.coord)
# 最右边列向外移动 1 列，最下边行向内移动 1 行，最左边列向内移动 1 列，最上边行向外移动 1 行
range.shrink(-1, 1, 1, -1)
print(range.coord)