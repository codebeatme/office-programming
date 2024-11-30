# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook

workbook = load_workbook('Goods.xlsx')
protection = workbook['Flowers'].protection
# 啟用對工作表 Flowers 的保護
protection.enable()
# 允許使用者在 Office 軟體中刪除列或欄
protection.deleteColumns = False
protection.deleteRows = False
# 設定密碼
protection.set_password('123')

workbook.save('Protection.xlsx')
