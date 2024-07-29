# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import open

workbook = open('Goods.xlsx')
worksheet = workbook['Flowers']
print(f'合併之前的 B2 為 {worksheet["B2"].value}')
# 合併範圍 A1:B2 之後，立即取消合併
worksheet.merge_cells('A1:B2')
print(f'合併之後的 B2 為 {worksheet["B2"].value}')
worksheet.unmerge_cells(start_row=1, start_column=1, end_row=2, end_column=2)

# 合併範圍 B2:C3，D4:H6
worksheet.merge_cells('B2:C3')
worksheet.merge_cells('D4:H6')
# 判斷儲存格 E3，D4 是否被合併了
print(f'儲存格 E3 被合併了嗎？{"E3" in worksheet.merged_cells}')
print(f'儲存格 D4 被合併了嗎？{"D4" in worksheet.merged_cells}')