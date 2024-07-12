from openpyxl.worksheet.cell_range import CellRange

# 创建区域 B2:D4，然后判断他与其他区域的关系
range = CellRange('B2:D4')
print(f'B2:D4 与 A1:B2 的交集为空？{range.isdisjoint(CellRange("A1:B2"))}')
print(f'B2:D4 与 F4:H5 的交集为空？{range.isdisjoint(CellRange("F4:H5"))}')

print(f'B2:D4 是否为 A1:D4 的子集？{range.issubset(CellRange("A1:D4"))}')
print(f'B2:D4 是否为 B2:D4 的子集？{range.issubset(CellRange("B2:D4"))}')

print(f'B2:D4 是否为 C3:C3 的超集？{range.issuperset(CellRange("C3:C3"))}')
print(f'B2:D4 是否为 B2:D4 的超集？{range.issuperset(CellRange("B2:D4"))}')
