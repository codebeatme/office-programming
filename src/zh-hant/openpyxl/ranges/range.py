from openpyxl.worksheet.cell_range import CellRange

# 建立範圍 B2:D4
range = CellRange('B2:D4')
# 顯示範圍的邊界
print(range.bounds)
print(f'最上方一列的儲存格的位置資訊 {range.top}')
print(f'最下方一列的儲存格的位置資訊 {range.bottom}')
print(f'最左邊一列的儲存格的位置資訊 {range.left}')
print(f'最右邊一列的儲存格的位置資訊 {range.right}')

# 建立範圍 C2:J4，工作表名稱為 SheetA
range = CellRange(min_col=3, min_row=2, max_col=10, max_row=4, title='SheetA')
# 顯示範圍的位址
print(range.coord)