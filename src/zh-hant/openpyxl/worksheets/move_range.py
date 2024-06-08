# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import open

workbook = open('Goods.xlsx')
worksheet = workbook['Pens']
# 移動包含資料和公式儲存格，並轉換公式
worksheet.move_range('A1:C3', 1, 1, True)
workbook.save('Move.xlsx')