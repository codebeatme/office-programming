from openpyxl.worksheet.cell_range import CellRange

# 创建区域 B2:D4，然后与其他区域进行计算
range = CellRange('B2:D4')
print(f'B2:D4 与 A1:B2 的交集：{range.intersection(CellRange("A1:B2"))}')

print(f'B2:D4 与 A1:D4 的并集：{range.union(CellRange("A1:D4"))}')
print(f'B2:D4 与 A1:A1 的并集：{range.union(CellRange("A1:A1"))}')

try:
    # ERROR 两个区域没有交集
    range.intersection(CellRange('A1:A1'))
except Exception as err:
    print(err)

try:
    # ERROR 目标区域的 title 有效并且与原区域不同
    range.union(CellRange('A1:A1', title='Other'))
except Exception as err:
    print(err)
