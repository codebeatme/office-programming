# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook

worksheet = load_workbook('Goods.xlsx')['Tables']
# 顯示工作表目前使用中的儲存格的位址
print(f'目前使用中的儲存格為 {worksheet.active_cell}')
