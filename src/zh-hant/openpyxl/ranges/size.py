from openpyxl.worksheet.cell_range import CellRange

# 建立區域 B2:D4
range = CellRange('B2:D4')
# 顯示區域大小
print(range.size)

# 最右邊欄向內移動 1 欄，最下邊列向外移動 1 列，最左邊欄向外移動 1 欄，最上邊列向內移動 1 列
range.expand(-1, 1, 1, -1)
print(range.coord)
# 最右邊欄向外移動 1 欄，最下邊列向內移動 1 列，最左邊欄向內移動 1 欄，最上邊列向外移動 1 列
range.shrink(-1, 1, 1, -1)
print(range.coord)