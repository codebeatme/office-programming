from openpyxl.worksheet.cell_range import CellRange

# 建立範圍 B2:D4，然後判斷他與其他範圍的關系
range = CellRange('B2:D4')
print(f'B2:D4 與 A1:B2 的交集為空？{range.isdisjoint(CellRange("A1:B2"))}')
print(f'B2:D4 與 F4:H5 的交集為空？{range.isdisjoint(CellRange("F4:H5"))}')

print(f'B2:D4 是否為 A1:D4 的子集？{range.issubset(CellRange("A1:D4"))}')
print(f'B2:D4 是否為 B2:D4 的子集？{range.issubset(CellRange("B2:D4"))}')

print(f'B2:D4 是否為 C3:C3 的超集？{range.issuperset(CellRange("C3:C3"))}')
print(f'B2:D4 是否為 B2:D4 的超集？{range.issuperset(CellRange("B2:D4"))}')
