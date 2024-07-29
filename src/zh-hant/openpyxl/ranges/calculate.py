from openpyxl.worksheet.cell_range import CellRange

# 建立範圍 B2:D4，然後與其他範圍進行計算
range = CellRange('B2:D4')
print(f'B2:D4 與 A1:B2 的交集：{range.intersection(CellRange("A1:B2"))}')

print(f'B2:D4 與 A1:D4 的並集：{range.union(CellRange("A1:D4"))}')
print(f'B2:D4 與 A1:A1 的並集：{range.union(CellRange("A1:A1"))}')

try:
    # ERROR 兩個範圍沒有交集
    range.intersection(CellRange('A1:A1'))
except Exception as err:
    print(err)

try:
    # ERROR 目標範圍的 title 有效並且與原範圍不同
    range.union(CellRange('A1:A1', title='Other'))
except Exception as err:
    print(err)
