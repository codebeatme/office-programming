from openpyxl.worksheet.cell_range import CellRange

# 创建区域 B2:D4
range = CellRange('B2:D4')
# 显示区域的边界
print(range.bounds)
print(f'最上方一行的单元格的位置信息 {range.top}')
print(f'最下方一行的单元格的位置信息 {range.bottom}')
print(f'最左边一行的单元格的位置信息 {range.left}')
print(f'最右边一行的单元格的位置信息 {range.right}')

# 创建区域 C2:J4，工作表名称为 SheetA
range = CellRange(min_col=3, min_row=2, max_col=10, max_row=4, title='SheetA')
# 显示区域的地址
print(range.coord)